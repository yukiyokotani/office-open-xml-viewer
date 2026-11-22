import { describe, expect, it } from 'vitest';
import type { ResolvedFontMetric } from '@silurus/ooxml-core';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { libreOfficeNormalLine, libreOfficeProducerLineMetric, openTypeNormalLine } from './layout/line-compatibility.js';
import { snapshotFontMetrics } from './layout/text.js';
import { projectedResourceMetric } from './line-breaker/font-metrics.js';
import type { LineLayout, ParagraphLayout } from './layout/types.js';
import type { BodyElement, DocParagraph, DocSettings, DocxDocumentModel, DocxTextRun, LineSpacing, SectionProps } from './types';

// LIBREOFFICE_FONT_LINE_BOX and LIBREOFFICE_PROPORTIONAL_LINE_PAGE_FIT: a
// model source that lays lines out as LibreOffice does sets
// `__libreOfficeLineBoxes`. The face below has LibreOffice's own geometry
// in proportion: an ascent of 0.8 em, a descent of 0.2 em and a 0.3 em line
// gap (the gap fixture font's hhea 800/-200/300), so Word's projection puts
// 1.1 em above the baseline and LibreOffice 0.8 em.

const GAP_FACE: ResolvedFontMetric = {
  family: 'Gap Face', requestedFamily: 'Gap Face', sourceIdentity: 'test-resource:gap-face',
  farEastCodePage: false,
  lineHeightRatio: 1.3, designAscentRatio: 1.1, designDescentRatio: 0.2,
  normalAscentRatio: 0.8, normalDescentRatio: 0.2, normalLineGapRatio: 0.3,
};
const LIBREOFFICE = { __libreOfficeLineBoxes: true } as unknown as DocSettings;

function linearCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = (): number => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => ({
      width: [...text].length * px() * 0.5,
      fontBoundingBoxAscent: px() * 0.8,
      fontBoundingBoxDescent: px() * 0.2,
      actualBoundingBoxAscent: px() * 0.8,
      actualBoundingBoxDescent: px() * 0.2,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

function section(pageHeight = 200): SectionProps {
  return {
    pageWidth: 200, pageHeight,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
  };
}

function run(text: string, fontSize = 12, extra: Partial<DocxTextRun> = {}): DocParagraph['runs'][number] {
  const textRun: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'Gap Face', isLink: false, background: null,
    vertAlign: null, hyperlink: null, ...extra,
  };
  return { type: 'text', ...textRun } as DocParagraph['runs'][number];
}
const BR = { type: 'break', breakType: 'line' } as DocParagraph['runs'][number];

function para(runs: DocParagraph['runs'], lineSpacing: LineSpacing | null = null, extra: object = {}): BodyElement {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing, numbering: null, tabStops: [],
    runs, defaultFontSize: 12, defaultFontFamily: 'Gap Face', widowControl: false, ...extra,
  } as unknown as BodyElement;
}

const auto = (value: number): LineSpacing => ({ rule: 'auto', value, explicit: true } as LineSpacing);

function pages(body: readonly BodyElement[], settings?: DocSettings, pageHeight?: number): LineLayout[][] {
  const layout = layoutBodyModel(body, section(pageHeight), linearCtx(), {}, settings, { 'gap face': GAP_FACE });
  return layout.pages.map((page) => page.layers.body
    .filter((node): node is ParagraphLayout => node.kind === 'paragraph')
    .flatMap((paragraph) => paragraph.lines));
}

describe('LIBREOFFICE_FONT_LINE_BOX', () => {
  it('reads a face\'s own normal line as LibreOffice selects it', () => {
    expect(libreOfficeNormalLine({ unitsPerEm: 1000, hhea: [880, -120, 500], win: [1015, 242] }))
      .toEqual({ normalAscentRatio: 0.88, normalDescentRatio: 0.12, normalLineGapRatio: 0.5 });
    // fsSelection USE_TYPO_METRICS: the OS/2 typo values instead.
    expect(libreOfficeNormalLine({ unitsPerEm: 1000, hhea: [880, -120, 500], typo: [750, -250, 200] }))
      .toEqual({ normalAscentRatio: 0.75, normalDescentRatio: 0.25, normalLineGapRatio: 0.2 });
    // An empty hhea falls back to usWin without a gap.
    expect(libreOfficeNormalLine({ unitsPerEm: 1000, hhea: [0, 0, 100], win: [900, 300] }))
      .toEqual({ normalAscentRatio: 0.9, normalDescentRatio: 0.3, normalLineGapRatio: 0 });
    expect(openTypeNormalLine({
      unitsPerEm: 1000, hheaAscent: 800, hheaDescent: -200, hheaLineGap: 300,
      typoAscent: 700, typoDescent: -300, typoLineGap: 0, useTypoMetrics: false, hasEastAsianCmap: false,
    })).toEqual({ normalAscentRatio: 0.8, normalDescentRatio: 0.2, normalLineGapRatio: 0.3 });
  });

  it('sizes a face for a LibreOffice producer from its own tables only', () => {
    // No OS/2 code-page class: Word's projection declines; the plain hhea
    // box (grid lines) and the normal line (LibreOffice's box) remain.
    expect(libreOfficeProducerLineMetric({
      unitsPerEm: 1000, hheaAscent: 800, hheaDescent: -200, hheaLineGap: 300,
      farEastCodePage: null, hasEastAsianCmap: true,
    }, [[0x20, 0x7e]])).toEqual({
      lineHeightRatio: 1.3, designAscentRatio: 1.1, designDescentRatio: 0.2,
      normalAscentRatio: 0.8, normalDescentRatio: 0.2, normalLineGapRatio: 0.3,
      unicodeRanges: [[0x20, 0x7e]], eastAsianLineHeightRatio: 1.3,
    });
  });

  it('keeps a Far East face\'s grid, LibreOffice and Word geometry apart', () => {
    const farEast = { unitsPerEm: 1000, hheaAscent: 800, hheaDescent: -200, hheaLineGap: 0, winAscent: 1000, winDescent: 300,
      farEastCodePage: true, hasEastAsianCmap: true };
    const metric = { family: 'FE', ...libreOfficeProducerLineMetric(farEast, undefined)! } as ResolvedFontMetric;
    // Grid lines: the plain hhea box (1 em), as catalogued faces give it.
    expect(projectedResourceMetric(metric, 'plain')!.lineHeightRatio).toBeCloseTo(1, 9);
    // Unsnapped lines: the face's own box, grown by 27% above under
    // MsWordCompGridMetrics.
    expect(projectedResourceMetric(metric, 'libreoffice')!.designAscentRatio).toBeCloseTo(0.8, 9);
    expect(projectedResourceMetric(metric, 'libreoffice-ms-word-grid')).toMatchObject({
      lineHeightRatio: expect.closeTo(1.27, 9), designAscentRatio: expect.closeTo(1.07, 9), designDescentRatio: expect.closeTo(0.2, 9) });
    // Vertical text keeps Word's projection: the same box.
    expect(projectedResourceMetric(metric, 'word')!.lineHeightRatio).toBeCloseTo(1.69, 9);
  });

  it('sizes a face with an empty hhea from usWin or typo values', () => {
    const empty = { unitsPerEm: 1000, hheaAscent: 0, hheaDescent: 0, hheaLineGap: 0, winAscent: 900, winDescent: 300,
      farEastCodePage: false, hasEastAsianCmap: false };
    expect(libreOfficeProducerLineMetric(empty, undefined)).toMatchObject({
      lineHeightRatio: 1.2, designAscentRatio: 0.9, designDescentRatio: 0.3, normalLineGapRatio: 0,
    });
    expect(libreOfficeProducerLineMetric({ ...empty, useTypoMetrics: true, typoAscent: 700, typoDescent: -300, typoLineGap: 200 },
      undefined)).toMatchObject({ lineHeightRatio: 1.2, designAscentRatio: 0.7, designDescentRatio: 0.5 });
  });

  it('keeps a measured negative gap as projection input, never an error', () => {
    const snapshot = snapshotFontMetrics({
      'tight face': { family: 'Tight Face', lineHeightRatio: 0.98, designAscentRatio: 0.78, designDescentRatio: 0.2,
        lineGapRatio: -0.02, normalAscentRatio: 0.8, normalDescentRatio: 0.2, normalLineGapRatio: -0.02 },
    });
    expect(snapshot['tight face']).toMatchObject({ normalLineGapRatio: -0.02, lineHeightRatio: 0.98 });
  });

  it('pins the text to the line top and puts the whole line gap below it', () => {
    const body = [para([run('One')]), para([run('Two')])];
    const [lines] = pages(body, LIBREOFFICE);
    // LibreOffice: baseline 0.8 em below the top, pitch 1.3 em.
    expect(lines!.map((line) => line.baselinePt)).toEqual([20 + 9.6, 20 + 15.6 + 9.6].map((v) => expect.closeTo(v, 6)));
    expect(lines!.map((line) => line.advancePt)).toEqual([15.6, 15.6].map((v) => expect.closeTo(v, 6)));
    // DOCX keeps Word's projection, the gap above the text.
    const [word] = pages(body);
    expect(word!.map((line) => line.baselinePt)).toEqual([20 + 13.2, 20 + 15.6 + 13.2].map((v) => expect.closeTo(v, 6)));
  });

  it('lets a tab take its text\'s box, not a synthetic em box', () => {
    // A face whose ascent (0.7 em) is below the synthetic 0.8 em tab box.
    const low: ResolvedFontMetric = { ...GAP_FACE, family: 'Low Face', requestedFamily: 'Low Face',
      sourceIdentity: 'test-resource:low-face', normalAscentRatio: 0.7, normalDescentRatio: 0.2, normalLineGapRatio: 0.4 };
    const layout = layoutBodyModel([para([{ ...run('a\tb'), fontFamily: 'Low Face' } as DocParagraph['runs'][number]])],
      section(), linearCtx(), {}, LIBREOFFICE, { 'low face': low });
    const line = (layout.pages[0]!.layers.body[0] as ParagraphLayout).lines[0]!;
    expect(line.advancePt).toBeCloseTo(15.6, 6);
    expect(line.baselinePt).toBeCloseTo(20 + 8.4, 6);
    // A line holding only a tab keeps its face's box too.
    const tabOnly = layoutBodyModel([para([{ ...run('\t'), fontFamily: 'Low Face' } as DocParagraph['runs'][number]])],
      section(), linearCtx(), {}, LIBREOFFICE, { 'low face': low });
    const tabLine = (tabOnly.pages[0]!.layers.body[0] as ParagraphLayout).lines[0]!;
    expect(tabLine.advancePt).toBeCloseTo(15.6, 6);
    expect(tabLine.baselinePt).toBeCloseTo(20 + 8.4, 6);
    // Without admitted face geometry the tab keeps the synthetic em box.
    const unknown = layoutBodyModel([para([{ ...run('\t'), fontFamily: 'Unknown Face' } as DocParagraph['runs'][number]])],
      section(), linearCtx(), {}, LIBREOFFICE, {});
    expect((unknown.pages[0]!.layers.body[0] as ParagraphLayout).lines[0]!.baselinePt).toBeCloseTo(20 + 9.6, 6);
  });

  it('takes a mixed line\'s largest ascent and largest descent with its gap', () => {
    const [lines] = pages([para([run('ab', 6), run('cd', 24)])], LIBREOFFICE);
    // 0.8 x 24 above, 0.5 x 24 below.
    expect(lines![0]!.baselinePt).toBeCloseTo(20 + 19.2, 6);
    expect(lines![0]!.advancePt).toBeCloseTo(31.2, 6);
  });

  it('keeps the text at the top of a proportional line', () => {
    const [lines] = pages([para([run('One'), BR, run('Two')], auto(2))], LIBREOFFICE);
    expect(lines!.map((line) => line.baselinePt)).toEqual([20 + 9.6, 20 + 31.2 + 9.6].map((v) => expect.closeTo(v, 6)));
  });

  it('keeps another producer\'s fixed line off the Word exact baseline (WORD_EXACT_LINE_BASELINE)', () => {
    const exact30 = { rule: 'exact', value: 30, explicit: true } as LineSpacing;
    // A Word document: four fifths of the 30pt line.
    const [word] = pages([para([run('Fixed')], exact30)]);
    expect(word![0]!.baselinePt).toBeCloseTo(20 + 24, 6);
    // A LibreOffice-producer line without its own fixed-line placement
    // (`__lineExtraAbove`, as for vertical ODT pages) keeps the centred box.
    const [other] = pages([para([run('Fixed')], exact30)], LIBREOFFICE);
    expect(other![0]!.advancePt).toBe(30);
    expect(other![0]!.baselinePt).toBeCloseTo(20 + (30 - 12) / 2 + 9.6, 6);
  });

  it('places a proportional line below 100% as a fixed line of its height', () => {
    const [lines] = pages([para([run('Low')], auto(0.8), { __lineExtraAbove: true })], LIBREOFFICE);
    expect(lines![0]!.advancePt).toBeCloseTo(12.48, 6);
    expect(lines![0]!.baselinePt).toBeCloseTo(20 + 12.48 * 4 / 5, 6);
  });
});

describe('LIBREOFFICE_PROPORTIONAL_LINE_PAGE_FIT', () => {
  // A 200 pt page with 20 pt margins: the text area ends at 180 pt. A
  // filler paragraph of fixed lines leaves `room` points above the bottom.
  const filler = (heightPt: number) => para([run('x')], { rule: 'exact', value: heightPt, explicit: true } as LineSpacing);
  const pageOf = (room: number, settings?: DocSettings) => {
    const body = [filler(160 - room), para([run('Probe'), BR, run('Next')], auto(2)), para([run('After')])];
    return pages(body, settings).map((lines) => lines.length);
  };

  it('keeps a line whose natural line fits, its extra space hanging below the margin', () => {
    // Probe takes its full 31.2 pt advance; Next needs its natural 15.6 pt.
    expect(pageOf(31.2 + 15.65, LIBREOFFICE)).toEqual([3, 1]);
    expect(pageOf(31.2 + 15.5, LIBREOFFICE)).toEqual([2, 2]);
  });

  it('gives an empty paragraph the allowance of a text line', () => {
    // The empty line (31.2 pt advance) stays when its 15.6 pt natural line
    // fits: After then starts page 2.
    const layout = (room: number) => layoutBodyModel([filler(160 - room), para([], auto(2)), para([run('After')])],
      section(), linearCtx(), {}, LIBREOFFICE, { 'gap face': GAP_FACE }).pages
      .map((page) => page.layers.body.filter((node) => node.kind === 'paragraph').length);
    expect(layout(15.65)).toEqual([2, 1]);
    expect(layout(15.5)).toEqual([1, 2]);
  });

  it('lets the extra space hang into a footnote area as into the margin', () => {
    // The note band (6 pt separator band and one 15.6 pt note line) starts
    // at 158.4 pt. Next carries the reference: it stays when its natural
    // line ends above the band, its extra 15.6 pt hanging into it, and the
    // body's flow charge ends at the band instead of claiming it.
    const reference = run('1', 12, { noteRef: { kind: 'footnote', id: '1' } });
    const layout = (room: number) => {
      const model = {
        body: [filler(160 - room), para([run('Probe'), BR, run('Next'), reference, BR, run('Third')], auto(2))],
        section: section(), headers: { default: null, first: null, even: null },
        footers: { default: null, first: null, even: null },
        footnotes: [{ id: '1', content: [para([run('', 12, { noteRef: { kind: 'footnote', id: '' } }), run('Note')])] }],
        endnotes: [], fontFamilyClasses: {}, settings: LIBREOFFICE,
      } as unknown as DocxDocumentModel;
      return layoutDocument(model, createLayoutServices(model, {
        measureContext: linearCtx(), localMetrics: { 'gap face': GAP_FACE },
      }), { currentDateMs: 0 });
    };
    // The probe starts at 111 pt: Probe's 31.2 pt and Next's natural
    // 15.6 pt end at 157.8 pt, 0.6 pt above the band.
    const fits = layout(69);
    const page = fits.pages[0]!;
    const probe = page.layers.body[1] as ParagraphLayout;
    const [band] = page.layers.notes;
    expect(probe.lines.length).toBe(2);
    expect(band!.flowBounds.yPt).toBeCloseTo(158.4, 6);
    const next = probe.lines[1]!;
    expect(next.bounds.yPt + 15.6).toBeLessThanOrEqual(158.4);
    expect(next.bounds.yPt + next.advancePt).toBeGreaterThan(158.4);
    expect(probe.flowBounds.yPt + probe.flowBounds.heightPt).toBeCloseTo(158.4, 6);
    // 1 pt less room: the natural line would end 0.4 pt inside the band,
    // so Next takes its note to the next page.
    const moved = layout(68);
    expect((moved.pages[0]!.layers.body[1] as ParagraphLayout).lines.length).toBe(1);
    expect(moved.pages[0]!.layers.notes).toEqual([]);
    expect(moved.pages[1]!.layers.notes.length).toBe(1);
  });

  it.each([
    { note: false, continuation: false },
    { note: true, continuation: false },
    { note: true, continuation: true },
  ])('cuts the space after before the hanging line (note $note, continuation $continuation)', ({ note, continuation }) => {
    // A completed two-line probe with 25 pt space after and contextual
    // spacing, followed by a paragraph of the same style. The region bottom
    // (the margin at 180 pt, or the note band at 158.4 pt) cuts the space
    // after first, so the following paragraph's collapse cannot rewind the
    // cursor into the hanging line: After starts the next page.
    const reference = note ? [run('1', 12, { noteRef: { kind: 'footnote', id: '1' } })] : [];
    const model = {
      body: [
        filler(91),
        para([run('Probe'), BR, run('Next'), ...reference], auto(2), { spaceAfter: 25, contextualSpacing: true, styleId: 'a' }),
        para([run('After')], null, { styleId: 'a' }),
      ],
      section: section(), headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      footnotes: [{ id: '1', content: [para([run('', 12, { noteRef: { kind: 'footnote', id: '' } }), run('Note')])] }],
      endnotes: [], fontFamilyClasses: {}, settings: LIBREOFFICE,
    } as unknown as DocxDocumentModel;
    const layout = layoutDocument(model, createLayoutServices(model, {
      measureContext: linearCtx(), localMetrics: { 'gap face': GAP_FACE }, allowFootnoteContinuation: continuation,
    }), { currentDateMs: 0 });
    const [first, second] = layout.pages;
    const probe = first!.layers.body[1] as ParagraphLayout;
    expect(first!.layers.body).toHaveLength(2);
    expect(probe.lines.map((line) => line.bounds.yPt)).toEqual([111, 142.2].map((v) => expect.closeTo(v, 6)));
    const bottomPt = note ? 158.4 : 180;
    const flowEndPt = probe.flowBounds.yPt + probe.flowBounds.heightPt;
    expect(flowEndPt).toBeGreaterThanOrEqual(142.2 + 15.6);
    expect(flowEndPt).toBeLessThanOrEqual(bottomPt + 1e-6);
    expect((second!.layers.body[0] as ParagraphLayout).lines[0]!.bounds.yPt).toBeCloseTo(20, 6);
  });

  it('keeps charging DOCX lines their whole advance', () => {
    expect(pageOf(31.2 + 15.65)).toEqual([2, 2]);
    expect(pageOf(31.2 + 31.25)).toEqual([3, 1]);
  });
});
