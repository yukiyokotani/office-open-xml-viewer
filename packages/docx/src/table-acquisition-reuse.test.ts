import { afterEach, describe, expect, it, vi } from 'vitest';
import * as tableLayoutModule from './layout/table.js';
import { retainedTableAcquisitionIsReusableAcrossPages } from './layout/table-acquisition.js';
import type { RetainedTableAcquisition } from './layout/table-acquisition.js';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import type { LayoutPage, TableLayoutInput } from './layout/types.js';
import type { TableFragmentLayout } from './layout/table-pagination.js';
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocTable,
  DocTableRow,
  DocxTextRun,
  SectionProps,
} from './types';

// Regression guard for the quadratic retained-table acquisition path: each
// flow region's `measureTable` used to re-acquire (re-measure and re-lay-out)
// the WHOLE table from row 0, so an R-row table paginating over P pages cost
// O(P·R) layout work. The acquisition is now reused while the inline extent
// is unchanged, so the canonical full-table layout runs exactly once.

type DocRun = DocParagraph['runs'][number];

function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, textAlign: 'left' as CanvasTextAlign,
    direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

function section(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: 200, pageHeight: 140,
    marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    ...overrides,
  };
}

function textRun(text: string, fontSize: number): DocRun {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  };
  return { type: 'text', ...run } as DocRun;
}

function para(text: string, fontSize = 20): CellElement {
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [textRun(text, fontSize)],
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics',
  };
  return { type: 'paragraph', ...p } as CellElement;
}

/** Single-column auto-height table whose rows wrap to several lines, so the
 *  table paginates over many pages with row splits at page boundaries. */
function wrappingTable(rowCount: number): BodyElement {
  const rows: DocTableRow[] = Array.from({ length: rowCount }, (_, index) => ({
    cells: [
      {
        content: [para(`rowmarker${index}rowmarker${index}`)],
        colSpan: 1,
        vMerge: null,
        borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
        background: null,
        vAlign: 'top',
        widthPt: null,
      },
    ],
    rowHeight: null,
    rowHeightRule: 'auto',
    isHeader: false,
  }));
  const t: DocTable = {
    colWidths: [160],
    rows,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
  };
  return { type: 'table', ...t } as BodyElement;
}

function pageRowSequence(
  pages: readonly LayoutPage[],
): readonly (readonly (readonly [number, number])[])[] {
  return pages.map((page) => page.layers.body
    .filter((node): node is TableFragmentLayout => node.kind === 'table')
    .flatMap((node) => node.rows
      .filter((row) => row.ownership === 'source')
      .map((row) => [row.logicalRowIndex, row.fragmentIndex] as const)));
}

afterEach(() => {
  vi.restoreAllMocks();
});

describe('retained table acquisition reuse across pages', () => {
  it('runs the canonical full-table layout once for a multi-page table', () => {
    const original = tableLayoutModule.layoutTable;
    let canonicalLayouts = 0;
    vi.spyOn(tableLayoutModule, 'layoutTable').mockImplementation(
      (input: TableLayoutInput, ...rest) => {
        // The retained acquisition lays the whole table out under its bare
        // flow-domain id; page-local fragments and occurrence probes carry
        // suffixed ids and are not acquisition work.
        if (input.id === 'table:0') canonicalLayouts += 1;
        return original(input, ...rest);
      },
    );

    const pages = layoutBodyModel([wrappingTable(20)], section(), makeCtx()).pages;

    expect(pages.length).toBeGreaterThan(2);
    expect(canonicalLayouts).toBe(1);
  });

  it('paginates a multi-page table identically when the acquisition is reused', () => {
    const pages = layoutBodyModel([wrappingTable(20)], section(), makeCtx()).pages;

    expect(pageRowSequence(pages)).toEqual([
      [[0, 0], [1, 0]],
      [[1, 1], [2, 0], [3, 0]],
      [[3, 1], [4, 0]],
      [[5, 0], [6, 0]],
      [[6, 1], [7, 0], [8, 0]],
      [[8, 1], [9, 0]],
      [[10, 0], [11, 0]],
      [[11, 1], [12, 0], [13, 0]],
      [[13, 1], [14, 0]],
      [[15, 0], [16, 0]],
      [[16, 1], [17, 0], [18, 0]],
      [[18, 1], [19, 0]],
    ]);
  });
});

describe('retainedTableAcquisitionIsReusableAcrossPages', () => {
  const fakeAcquisition = (
    block: object,
    nested?: RetainedTableAcquisition,
  ): RetainedTableAcquisition => ({
    input: {
      rows: [{ cells: [{ blocks: [block] }] }],
    },
    nestedById: nested ? Object.freeze({ nested }) : Object.freeze({}),
  }) as unknown as RetainedTableAcquisition;

  const paragraphBlock = (extra: object = {}) => ({
    layout: { kind: 'paragraph', lines: [], drawings: [], textBoxes: [], ...extra },
  });

  it('accepts an acquisition of plain paragraph rows', () => {
    expect(retainedTableAcquisitionIsReusableAcrossPages(
      fakeAcquisition(paragraphBlock()),
    )).toBe(true);
  });

  it('rejects an acquisition carrying page-dependent field blocks', () => {
    expect(retainedTableAcquisitionIsReusableAcrossPages(
      fakeAcquisition({ ...paragraphBlock(), pageDependent: true }),
    )).toBe(false);
  });

  it('rejects an acquisition carrying anchored drawings', () => {
    const anchored = paragraphBlock({
      drawings: [{ kind: 'drawing', anchorLayer: { occurrenceId: 'a' } }],
    });
    expect(retainedTableAcquisitionIsReusableAcrossPages(
      fakeAcquisition(anchored),
    )).toBe(false);
  });

  it('rejects page-dependent content nested in a text box', () => {
    const pageDependentTextBox = paragraphBlock({
      textBoxes: [{
        kind: 'textbox',
        story: {
          blocks: [{
            kind: 'paragraph',
            lines: [{ placements: [{ kind: 'text', dependency: 'page' }] }],
            drawings: [],
            textBoxes: [],
          }],
        },
      }],
    });
    expect(retainedTableAcquisitionIsReusableAcrossPages(
      fakeAcquisition(pageDependentTextBox),
    )).toBe(false);
  });

  it('accepts page-invariant content nested in a text box', () => {
    const plainTextBox = paragraphBlock({
      textBoxes: [{
        kind: 'textbox',
        story: {
          blocks: [{
            kind: 'paragraph',
            lines: [{ placements: [{ kind: 'text' }] }],
            drawings: [],
            textBoxes: [],
          }],
        },
      }],
    });
    expect(retainedTableAcquisitionIsReusableAcrossPages(
      fakeAcquisition(plainTextBox),
    )).toBe(true);
  });

  it('rejects anchored drawings nested in a text box', () => {
    const anchoredTextBox = paragraphBlock({
      textBoxes: [{
        kind: 'textbox',
        story: {
          blocks: [{
            kind: 'paragraph',
            lines: [],
            drawings: [{ kind: 'drawing', anchorLayer: { occurrenceId: 'nested' } }],
            textBoxes: [],
          }],
        },
      }],
    });
    expect(retainedTableAcquisitionIsReusableAcrossPages(
      fakeAcquisition(anchoredTextBox),
    )).toBe(false);
  });

  it('rejects when a nested table is not reusable', () => {
    const nested = fakeAcquisition({ ...paragraphBlock(), pageDependent: true });
    expect(retainedTableAcquisitionIsReusableAcrossPages(
      fakeAcquisition(paragraphBlock(), nested),
    )).toBe(false);
  });
});
