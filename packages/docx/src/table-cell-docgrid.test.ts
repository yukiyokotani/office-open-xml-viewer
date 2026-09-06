import { describe, expect, it } from 'vitest';
import {
  layoutBodyModel,
  layoutBodyTableRowAdvances,
} from './test-support/document-layout.test-support.js';
import type {
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
  SectionProps,
} from './types.js';

const RESOLVED_EA_FAMILY = 'Arbitrary Resolved EA';
const RESOLVED_EA_RATIO = 3269 / 2048;
const RESOLVED_EA_METRICS = {
  'arbitrary resolved ea': {
    family: RESOLVED_EA_FAMILY,
    eastAsianLineHeightRatio: RESOLVED_EA_RATIO,
  },
};

function makeMeasureContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  // The mock glyph box is a flat 1.0×em (ascent 0.8 + descent 0.2 of the
  // CURRENT font size). No run family is in the core metric table, so every
  // height below is exactly the tallest run's synthetic box — any growth
  // beyond it must come from the grid cell rounding under test.
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() {
      return font;
    },
    set font(value: string) {
      font = value;
    },
    letterSpacing: '0px',
    measureText: (text: string) => ({
      width: [...text].length * 10,
      fontBoundingBoxAscent: px() * 0.8,
      fontBoundingBoxDescent: px() * 0.2,
      actualBoundingBoxAscent: px() * 0.8,
      actualBoundingBoxDescent: px() * 0.2,
    }) as TextMetrics,
    save() {},
    restore() {},
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

function section(docGridType: 'lines' | 'snapToChars' = 'lines'): SectionProps {
  return {
    pageWidth: 200,
    pageHeight: 300,
    marginTop: 20,
    marginRight: 20,
    marginBottom: 20,
    marginLeft: 20,
    headerDistance: 10,
    footerDistance: 10,
    titlePage: false,
    evenAndOddHeaders: false,
    docGridType,
    docGridLinePitch: 20,
  };
}

function paragraph(): DocParagraph {
  return {
    alignment: 'left',
    indentLeft: 0,
    indentRight: 0,
    indentFirst: 0,
    spaceBefore: 0,
    spaceAfter: 0,
    lineSpacing: null,
    numbering: null,
    tabStops: [],
    runs: [{
      type: 'text',
      text: 'あ',
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      fontSize: 10,
      color: null,
      fontFamily: 'serif',
      fontFamilyEastAsia: 'serif',
      isLink: false,
      background: null,
      vertAlign: null,
      hyperlink: null,
    }],
    defaultFontSize: 10,
    defaultFontFamily: 'serif',
    widowControl: false,
  } as unknown as DocParagraph;
}

/** A paragraph with caller-supplied runs (fontSize in pt; text runs get the
 *  same inert defaults as {@link paragraph}). */
function paragraphWithRuns(runs: Record<string, unknown>[]): DocParagraph {
  const base = paragraph() as unknown as { runs: unknown[] };
  base.runs = runs.map((r) => (
    r.type === 'break'
      ? r
      : {
          type: 'text',
          bold: false, italic: false, underline: false, strikethrough: false,
          color: null, fontFamily: 'serif', fontFamilyEastAsia: 'serif',
          isLink: false, background: null, vertAlign: null, hyperlink: null,
          ...r,
        }
  ));
  return base as unknown as DocParagraph;
}

function cellWith(para: DocParagraph): DocTableCell {
  const c = cell() as unknown as { content: unknown[] };
  c.content = [{ type: 'paragraph', ...para }];
  return c as unknown as DocTableCell;
}

function rowWith(para: DocParagraph): DocTableRow {
  const r = row() as unknown as { cells: unknown[] };
  r.cells = [cellWith(para)];
  return r as unknown as DocTableRow;
}

function cell(): DocTableCell {
  return {
    content: [{ type: 'paragraph', ...paragraph() }],
    colSpan: 1,
    vMerge: null,
    borders: {
      top: null,
      bottom: null,
      left: null,
      right: null,
      insideH: null,
      insideV: null,
    },
    background: null,
    vAlign: 'top',
    widthPt: 100,
    marginTop: 0,
    marginBottom: 0,
    marginLeft: 0,
    marginRight: 0,
  } as unknown as DocTableCell;
}

function row(): DocTableRow {
  return {
    cells: [cell()],
    rowHeight: null,
    rowHeightRule: 'auto',
    isHeader: false,
  } as unknown as DocTableRow;
}

function table(): DocTable {
  return {
    colWidths: [100],
    rows: [],
    borders: {
      top: null,
      bottom: null,
      left: null,
      right: null,
      insideH: null,
      insideV: null,
    },
    cellMarginTop: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 0,
    cellMarginRight: 0,
    jc: 'left',
  } as unknown as DocTable;
}

function retainedAdvance(
  tableRow: DocTableRow,
  adjustLineHeightInTable: boolean,
  docGridType: 'lines' | 'snapToChars' = 'lines',
): number {
  const t = table();
  const advance = layoutBodyTableRowAdvances(
    { ...t, rows: [tableRow] },
    section(docGridType),
    makeMeasureContext(),
    { adjustLineHeightInTable },
    RESOLVED_EA_METRICS,
  )[0];
  if (advance === undefined) throw new Error('Canonical table omitted the row');
  return advance;
}

describe('table-cell line grid compatibility', () => {
  it('keeps cell text at natural height when compatibility is disabled', () => {
    expect(retainedAdvance(row(), false)).toBe(10);
  });

  it('applies the section line pitch when compatibility is enabled', () => {
    expect(retainedAdvance(row(), true)).toBe(20);
  });

  it('gates the line axis of snapToChars by the same compatibility setting', () => {
    expect(retainedAdvance(row(), false, 'snapToChars')).toBe(10);
    expect(retainedAdvance(row(), true, 'snapToChars')).toBe(20);
  });
});

// The docGrid line-cell count (docGridLineCells) is derived from the line's
// resolved single-line height — the TALLEST run's box governs the line
// (ECMA-376 §17.3.1.33; the grid reserves the whole cells that CONTAIN that
// box, §17.6.5 / issue #1013 sample-58 adjudication). These integration cases
// pin two call-path properties the pure lineBoxHeight tests cannot see: which
// line height the layout hands over, and which script gate each LINE uses. The
// grid is active in a cell via adjustLineHeightInTable (§17.15.3.1); pitch =
// 20 pt; the mock font box is a flat 1.0×em, so every height difference below
// comes from the line-height / script routing alone.
describe('docGrid line-cell integration through the cell measure path', () => {
  it('retains atLeast-zero empty-mark row and horizontal-border boundaries', () => {
    const para = paragraphWithRuns([]);
    para.defaultFontSize = 10;
    para.defaultFontFamily = RESOLVED_EA_FAMILY;
    para.defaultFontFamilyEastAsia = RESOLVED_EA_FAMILY;
    para.lineSpacing = { value: 0, rule: 'atLeast', explicit: true };
    const source = table();
    const horizontal = { width: 2, color: '#000000', style: 'single' };
    source.rows = [rowWith(para)];
    source.borders = { ...source.borders, top: horizontal, bottom: horizontal };
    const gridSection = { ...section(), docGridLinePitch: 14.55 };

    const layout = layoutBodyModel(
      [{ type: 'table', ...source }],
      gridSection,
      makeMeasureContext(),
      {},
      { adjustLineHeightInTable: true, useFeLayout: true },
      RESOLVED_EA_METRICS,
    );
    const retained = layout.pages[0]?.layers.body[0];
    if (retained?.kind !== 'table') throw new Error('Canonical layout omitted the table');
    const retainedRow = retained.rows[0];
    if (!retainedRow) throw new Error('Canonical table omitted the row');
    const top = retained.borders.find((border) => border.edge === 'top');
    const bottom = retained.borders.find((border) => border.edge === 'bottom');
    if (!top || !bottom) throw new Error('Canonical table omitted horizontal borders');

    const designAdvancePt = 10 * RESOLVED_EA_RATIO;
    expect(retainedRow.contentHeightPt).toBeCloseTo(designAdvancePt, 12);
    expect(retainedRow.advancePt).toBeCloseTo(designAdvancePt + 2, 12);
    expect(retained.advancePt).toBeCloseTo(retainedRow.advancePt, 12);
    expect(top.from.yPt).toBeCloseTo(retained.flowBounds.yPt, 12);
    expect(bottom.from.yPt).toBeCloseTo(
      retained.flowBounds.yPt + retained.advancePt,
      12,
    );
  });

  it('matches observed Word spacing for explicit atLeast lines in a table cell', () => {
    const para = paragraphWithRuns([
      {
        text: 'あ',
        fontSize: 14,
        fontFamily: RESOLVED_EA_FAMILY,
        fontFamilyEastAsia: RESOLVED_EA_FAMILY,
      },
      { type: 'break', breakType: 'line' },
      {
        text: 'い',
        fontSize: 10,
        fontFamily: RESOLVED_EA_FAMILY,
        fontFamilyEastAsia: RESOLVED_EA_FAMILY,
      },
    ]);
    para.lineSpacing = { value: 0, rule: 'atLeast', explicit: true };

    // Observed Windows Word output keeps the first explicit-atLeast line at
    // The resolved resource's raw 14pt design height (3269/2048 em, slightly
    // over the 20pt pitch) is followed by one ordinary 10pt grid line.
    expect(retainedAdvance(rowWith(para), true))
      .toBeCloseTo(14 * RESOLVED_EA_RATIO + 20, 12);
  });

  it('a manual line break in a SMALLER run does not shrink the line height (tallest governs)', () => {
    // Line 1: 'あ' at 24 pt + 'い' at 10 pt, terminated by a <w:br> whose
    // nearby size resolves to 10 pt (§17.3.3.1; findNearbyFontSize looks at the
    // preceding run). Line 2: 'う' at 10 pt. The break must NOT overwrite the
    // line's tallest box (24 px → ceil(24/20) = 2 cells = 40) with its own
    // 10 pt box (1 cell = 20).
    const para = paragraphWithRuns([
      { text: 'あ', fontSize: 24 },
      { text: 'い', fontSize: 10 },
      { type: 'break', breakType: 'line' },
      { text: 'う', fontSize: 10 },
    ]);
    expect(retainedAdvance(rowWith(para), true))
      .toBe(60); // 40 (2 cells) + 20 (1 cell)
  });

  it('the East Asian cell rounding is gated per LINE, not per paragraph', () => {
    // Line 1: CJK 10 pt → 1 cell (20). Line 2: Latin-only 'Hello' at 22 pt —
    // Word does not cell-round Latin lines; they keep their natural height
    // above a one-cell floor (mock natural = 22 px > floor 20 → 22), NOT the
    // cell count ceil(22/20) = 2 cells = 40 that a paragraph-level East Asian
    // flag would apply.
    const para = paragraphWithRuns([
      { text: 'あ', fontSize: 10 },
      { type: 'break', breakType: 'line' },
      { text: 'Hello', fontSize: 22 },
    ]);
    expect(retainedAdvance(rowWith(para), true))
      .toBe(42); // 20 (CJK, 1 cell) + 22 (Latin, natural above the one-cell floor)
  });
});
