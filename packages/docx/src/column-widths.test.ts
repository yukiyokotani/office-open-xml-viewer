import { describe, it, expect } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { resolveColumnWidths } from './test-support/renderer-internals.test-support.js';
import { buildSegments, layoutLines } from './line-layout.js';
import { bodyAcquisitionInputProjections } from './parser-model.js';
import { canvasFontString } from '@silurus/ooxml-core';
import {
  resolveDocumentLayoutSettings,
  resolveSectionLayoutContext,
} from './layout-context.js';
import type {
  DocParagraph,
  DocTable,
  DocTableRow,
  DocTableCell,
  DocxDocumentModel,
} from './types.js';

// State type taken from resolveColumnWidths' signature so the test does not
// depend on the acquisition-state type being re-exported (mirrors table-row-height.test).
type ColState = Parameters<typeof resolveColumnWidths>[2];

// ECMA-376 §17.4.48 (`<w:tblGrid>`) supplies the INITIAL shared grid;
// §17.18.87 then applies tblW, wBefore/wAfter and tcW preferences. A saved
// grid is not evidence that a producer already applied those constraints.
//
// Empty cells carry no content, so the min-content floor is 0. The projection
// capability remains required because table-format facts are acquired before
// the empty-content fast path.

function normalizedColumnState(state: Record<string, unknown>): ColState {
  const document = {
    section: {
      pageWidth: 300, pageHeight: 400,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 10, footerDistance: 10,
      titlePage: false, evenAndOddHeaders: false,
      sectionStart: 'nextPage', columns: null,
    },
    body: [], headers: {}, footers: {},
  } as unknown as DocxDocumentModel;
  const layoutSettings = resolveDocumentLayoutSettings(document);
  return {
    ...state,
    layoutSettings,
    sectionLayout: resolveSectionLayoutContext(layoutSettings, document.section),
  } as unknown as ColState;
}

const EMPTY_STATE = normalizedColumnState({
  acquisitionInputs: bodyAcquisitionInputProjections,
});

function cell(colSpan: number, widthPct: number | null, widthPt: number | null = null): DocTableCell {
  return {
    content: [],
    colSpan,
    vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background: null,
    vAlign: 'top',
    widthPt,
    widthPct,
    marginTop: 0,
    marginBottom: 0,
    marginLeft: 0,
    marginRight: 0,
  } as unknown as DocTableCell;
}

function row(cells: DocTableCell[]): DocTableRow {
  return { cells, rowHeight: null, rowHeightRule: 'auto', isHeader: false } as unknown as DocTableRow;
}

/** sample-3-shaped table: a fixed-preferred-width table (`tblW=pct`, the
 *  whole-table preferred width Word baked its auto-fit into) with an unequal
 *  2-column grid whose single-column cells each prefer ~50% via `tcW=pct`. */
function table(rows: DocTableRow[], width: Partial<DocTable> = { widthPct: 5000 }): DocTable {
  return {
    colWidths: [70, 30], // grid: deliberately UNEQUAL
    rows,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 0,
    cellMarginRight: 0,
    jc: 'left',
    ...width,
    // layout omitted ⇒ autofit (the branch under test).
  } as unknown as DocTable;
}

describe('resolveColumnWidths — fixed preferences are applied over the initial tblGrid', () => {
  it('applies single-column tcW=pct constraints over an unequal initial grid', () => {
    // Two first-row single-column cells each prefer 50% of the final table.
    const t = table([
      row([cell(1, 2500), cell(1, 2500)]),
      row([cell(1, 2500), cell(1, 2500)]),
    ]);
    const w = resolveColumnWidths(t, 100, EMPTY_STATE);
    expect(w[0]).toBeCloseTo(50, 5);
    expect(w[1]).toBeCloseTo(50, 5);
  });

  it('a gridSpan cell whose tcW exceeds its grid span does not widen the columns', () => {
    // A 2-col-spanning cell prefers 100% (5000pct) — wider than the grid's full
    // 100 pt. The grid still governs: columns stay 70/30, total stays 100.
    const t = table([row([cell(2, 5000)])]);
    const w = resolveColumnWidths(t, 100, EMPTY_STATE);
    expect(w[0]).toBeCloseTo(70, 5);
    expect(w[1]).toBeCloseTo(30, 5);
  });

  it('resolves percentage tcW against the smaller preferred table width', () => {
    const t = table([row([cell(1, 2500), cell(1, 2500)])]);
    const w = resolveColumnWidths(t, 50, EMPTY_STATE);
    expect(w[0]).toBeCloseTo(25, 5);
    expect(w[1]).toBeCloseTo(25, 5);
  });
});

// ECMA-376 §17.4.63 (`<w:tblW>`) / §17.18.87 (ST_TblWidth "auto"). A
// tblW=auto table ("AutoFit to Contents") has NO preferred table width: Word
// sizes columns from content + per-cell `<w:tcW>`, and the saved `<w:gridCol>`
// is the style/layout default (frequently the full text column), NOT a baked
// auto-fit result. So for tblW=auto the grid must NOT drive the widths — tcW /
// content does. (sample-7's cover tables carry gridCol=full-page yet a 100 pt
// tcW; trusting the grid made them full-width and defeated their own w:jc
// right/left placement.) Contrast the tblW=dxa/pct case above, where the grid
// IS Word's baked layout and stays authoritative (sample-3).
describe('resolveColumnWidths — a tblW=auto table sizes to tcW/content, ignoring the stale full-width grid', () => {
  const autofit = (rows: DocTableRow[], colWidths: number[]): DocTable =>
    ({
      ...table(rows, { widthPt: undefined, widthPct: undefined }),
      colWidths,
    }) as unknown as DocTable;

  it('a single-column tblW=auto table settles at the cell tcW (not the full-page grid)', () => {
    // sample-7 leaders table: gridCol = full content width (415 pt-ish, here
    // 415), but the cell prefers tcW=100 pt. AutoFit-to-Contents ⇒ column = 100.
    const t = autofit([row([cell(1, null, 100)]), row([cell(1, null, 100)])], [415]);
    const w = resolveColumnWidths(t, 415, EMPTY_STATE);
    expect(w[0]).toBeCloseTo(100, 5);
  });

  it('a two-column tblW=auto table uses each cell tcW, leaving the table narrower than the page', () => {
    // sample-7 Arabic table: gridCol=[207.65,207.65] (full width), but the cells
    // prefer tcW col0=120, col1=20. The resolved table is 140 wide (< 415), so
    // its w:jc can place it at the right margin. The grid (415) is ignored.
    const t = autofit(
      [row([cell(1, null, 120), cell(1, null, 20)])],
      [207.65, 207.65],
    );
    const w = resolveColumnWidths(t, 415, EMPTY_STATE);
    expect(w[0]).toBeCloseTo(120, 5);
    expect(w[1]).toBeCloseTo(20, 5);
    expect(w[0] + w[1]).toBeLessThan(415);
  });

  it('uses the same registered highAnsi substitute route as ordinary line layout', () => {
    let font = '';
    const measured: Array<{ text: string; font: string }> = [];
    const ctx = {
      get font() { return font; },
      set font(value: string) { font = value; },
      letterSpacing: '0px',
      fontKerning: 'auto' as CanvasFontKerning,
      measureText(text: string) {
        measured.push({ text, font });
        return {
          width: [...text].length * 10,
          fontBoundingBoxAscent: 8,
          fontBoundingBoxDescent: 2,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2,
        } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;
    const doc = {
      section: {}, body: [], headers: {}, footers: {}, majorFont: 'Calibri',
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(doc, {
      providerRoutes: {
        calibri: { family: 'Carlito', source: 'substitute' },
      },
      measureContext: ctx,
    });
    const run = {
      type: 'text', text: 'é', fontSize: 10, fontFamily: 'Legacy ASCII',
      fontFamilyHighAnsi: 'Calibri', bold: false, italic: false, underline: false,
      strikethrough: false, color: null, isLink: false, background: null,
      vertAlign: null, hyperlink: null,
      fontSlots: {
        direct: { ascii: 'Legacy ASCII', highAnsi: 'Calibri' },
        theme: {},
        themePresent: { ascii: false, highAnsi: false, eastAsia: false, complexScript: false },
      },
    } as DocParagraph['runs'][number];
    const paragraph = {
      alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
      tabStops: [], runs: [run], widowControl: false,
    } as DocParagraph;
    const textCell = {
      ...cell(1, null, 0), content: [{ type: 'paragraph', ...paragraph }],
    } as unknown as DocTableCell;
    const t = autofit([row([textCell])], [200]);
    const state = normalizedColumnState({
      ctx, fontFamilyClasses: {}, layoutServices: services, pageWidth: 300,
      acquisitionInputs: bodyAcquisitionInputProjections,
    });

    resolveColumnWidths(t, 300, state);
    const autoFitFonts = measured
      .filter((entry) => entry.text === 'é')
      .map((entry) => entry.font);
    measured.length = 0;
    const segments = buildSegments([run], {
      pageIndex: 0, totalPages: 1, layoutServices: services,
    });
    layoutLines(ctx, segments, 300, 0, 1);
    const expectedShape = services.text.shape({
      text: 'é', fontSizePt: 10,
      fonts: { ascii: 'Legacy ASCII', highAnsi: 'Calibri' },
    });
    const expectedFont = canvasFontString(
      expectedShape.spans[0]!.fontRoute, 10, 400, 'normal',
    );

    expect(expectedShape.spans[0]!.font).toMatchObject({
      source: 'substitute', resolvedFamily: 'Carlito',
    });
    expect(segments[0]).toMatchObject({ fontRoute: expectedShape.spans[0]!.fontRoute });
    expect(autoFitFonts).toContain(expectedFont);
    // Ordinary line layout reuses the immutable shape acquired by auto-fit.
    // The retained segment route, rather than another native measurement call,
    // proves that both consumers use the same registered substitute.
    expect(measured.filter((entry) => entry.text === 'é')).toEqual([]);
  });

  it('feeds the unwrapped paragraph width into the autofit maximum-content step', () => {
    const ctx = {
      font: '', letterSpacing: '0px', fontKerning: 'auto' as CanvasFontKerning,
      measureText(text: string) {
        return {
          width: [...text].length * 10,
          fontBoundingBoxAscent: 8,
          fontBoundingBoxDescent: 2,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2,
        } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;
    const doc = { section: {}, body: [], headers: {}, footers: {} } as unknown as DocxDocumentModel;
    const services = createLayoutServices(doc, { measureContext: ctx });
    const paragraph = (text: string) => ({
      alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
      tabStops: [], widowControl: false,
      runs: [{
        type: 'text', text, fontSize: 10, fontFamily: 'Synthetic Serif',
        fontFamilyEastAsia: '', bold: false, italic: false, underline: false,
        strikethrough: false, color: null, isLink: false, background: null,
        vertAlign: null, hyperlink: null,
      }],
    }) as DocParagraph;
    const contentCell = (text: string) => ({
      ...cell(1, null, null),
      content: [{ type: 'paragraph', ...paragraph(text) }],
    }) as unknown as DocTableCell;
    const t = autofit([row([contentCell('aa aa'), contentCell('b')])], [0, 100]);
    const state = normalizedColumnState({
      ctx, fontFamilyClasses: {}, layoutServices: services, pageWidth: 100,
      pageH: 200, pageIndex: 0, totalPages: 1,
      acquisitionInputs: bodyAcquisitionInputProjections,
    });

    // min-content is [20, 10], but the first cell's no-wrap maximum is 50.
    // Autofit first reclaims the second track's slack toward that maximum.
    expect(resolveColumnWidths(t, 100, state)).toEqual([50, 50]);
  });

  it('sums each retained route across a differently formatted no-break pair', () => {
    let font = '';
    const measured: Array<{ text: string; font: string }> = [];
    const ctx = {
      get font() { return font; },
      set font(value: string) { font = value; },
      letterSpacing: '0px',
      fontKerning: 'auto' as CanvasFontKerning,
      measureText(text: string) {
        measured.push({ text, font });
        const perScalar = font.includes('Carlito') ? 13 : font.includes('Caladea') ? 29 : 5;
        return {
          width: [...text].length * perScalar,
          fontBoundingBoxAscent: 8,
          fontBoundingBoxDescent: 2,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2,
        } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;
    const doc = {
      section: {}, body: [], headers: {}, footers: {}, majorFont: 'Calibri',
    } as unknown as DocxDocumentModel;
    const services = createLayoutServices(doc, {
      providerRoutes: {
        calibri: { family: 'Carlito', source: 'substitute' },
        cambria: { family: 'Caladea', source: 'substitute' },
      },
      measureContext: ctx,
    });
    const textRun = (text: string, family: string): DocParagraph['runs'][number] => ({
      type: 'text', text, fontSize: 10, fontFamily: family,
      fontFamilyHighAnsi: family, bold: false, italic: false, underline: false,
      strikethrough: false, color: null, isLink: false, background: null,
      vertAlign: null, hyperlink: null,
      fontSlots: {
        direct: { ascii: family, highAnsi: family },
        theme: {},
        themePresent: { ascii: false, highAnsi: false, eastAsia: false, complexScript: false },
      },
    }) as DocParagraph['runs'][number];
    const paragraph = {
      alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
      tabStops: [], widowControl: false,
      // UAX #14 LB14 forbids a break between OP and AL. The formatting seam is
      // not a license to route the following scalar through the opening mark's
      // font: each piece keeps its own TextLayoutService request.
      runs: [textRun('(', 'Carlito'), textRun('A', 'Caladea')],
    } as DocParagraph;
    const textCell = {
      ...cell(1, null, 0), content: [{ type: 'paragraph', ...paragraph }],
    } as unknown as DocTableCell;
    const t = autofit([row([textCell])], [200]);
    const state = normalizedColumnState({
      ctx, fontFamilyClasses: {}, layoutServices: services, pageWidth: 300,
      acquisitionInputs: bodyAcquisitionInputProjections,
    });

    const segments = buildSegments(paragraph.runs, {
      pageIndex: 0, totalPages: 1, layoutServices: services,
    });
    expect(segments).toMatchObject([
      { text: '(', fontFamily: 'Carlito' },
      { text: 'A', fontFamily: 'Caladea', joinPrev: true },
    ]);
    expect(resolveColumnWidths(t, 300, state)[0]).toBe(42);
    expect(measured).toEqual(expect.arrayContaining([
      expect.objectContaining({ text: '(', font: expect.stringContaining('Carlito') }),
      expect.objectContaining({ text: 'A', font: expect.stringContaining('Caladea') }),
    ]));
    expect(measured.some((entry) => entry.text === '(A')).toBe(false);
  });

  it('keeps a CJK grapheme cluster atomic when deriving the auto-fit minimum', () => {
    const ctx = {
      font: '', letterSpacing: '0px', fontKerning: 'auto' as CanvasFontKerning,
      measureText(text: string) {
        return {
          width: [...text].length * 10,
          fontBoundingBoxAscent: 8,
          fontBoundingBoxDescent: 2,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2,
        } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;
    const doc = { section: {}, body: [], headers: {}, footers: {} } as unknown as DocxDocumentModel;
    const services = createLayoutServices(doc, { measureContext: ctx });
    const paragraph = {
      alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
      tabStops: [], widowControl: false,
      runs: [{
        type: 'text', text: 'か\u3099国', fontSize: 10, fontFamily: 'EA Face',
        fontFamilyEastAsia: 'EA Face', bold: false, italic: false, underline: false,
        strikethrough: false, color: null, isLink: false, background: null,
        vertAlign: null, hyperlink: null,
      }],
    } as DocParagraph;
    const textCell = {
      ...cell(1, null, 0), content: [{ type: 'paragraph', ...paragraph }],
    } as unknown as DocTableCell;
    const t = autofit([row([textCell])], [200]);
    const state = normalizedColumnState({
      ctx, fontFamilyClasses: {}, layoutServices: services, pageWidth: 300,
      acquisitionInputs: bodyAcquisitionInputProjections,
    });

    expect(resolveColumnWidths(t, 300, state)[0]).toBe(20);
  });

  it('keeps a kinsoku-prohibited CJK boundary inside one auto-fit minimum atom', () => {
    const ctx = {
      font: '', letterSpacing: '0px', fontKerning: 'auto' as CanvasFontKerning,
      measureText(text: string) {
        return {
          width: [...text].length * 10,
          fontBoundingBoxAscent: 8,
          fontBoundingBoxDescent: 2,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2,
        } as TextMetrics;
      },
    } as unknown as CanvasRenderingContext2D;
    const doc = { section: {}, body: [], headers: {}, footers: {} } as unknown as DocxDocumentModel;
    const services = createLayoutServices(doc, { measureContext: ctx });
    const paragraph = {
      alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
      tabStops: [], widowControl: false,
      runs: [{
        type: 'text', text: '（国', fontSize: 10, fontFamily: 'EA Face',
        fontFamilyEastAsia: 'EA Face', bold: false, italic: false, underline: false,
        strikethrough: false, color: null, isLink: false, background: null,
        vertAlign: null, hyperlink: null,
      }],
    } as DocParagraph;
    const textCell = {
      ...cell(1, null, 0), content: [{ type: 'paragraph', ...paragraph }],
    } as unknown as DocTableCell;
    const t = autofit([row([textCell])], [200]);
    const state = normalizedColumnState({
      ctx, fontFamilyClasses: {}, layoutServices: services, pageWidth: 300,
      acquisitionInputs: bodyAcquisitionInputProjections,
    });

    expect(resolveColumnWidths(t, 300, state)[0]).toBe(20);
  });
});
