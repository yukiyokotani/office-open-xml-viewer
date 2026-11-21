import { describe, it, expect } from 'vitest';
import { layoutBodyTableRowAdvances } from './test-support/document-layout.test-support.js';
import type {
  DocTable,
  DocTableRow,
  DocTableCell,
  DocParagraph,
  SectionProps,
} from './types.js';

// ECMA-376 §17.3.1.12 (`<w:ind>`) + §17.4.80 (auto row height): a table cell's
// auto row height is the tallest cell's measured CONTENT height. The content is
// laid out inside `cellWidth − cell margins − the paragraph's own left/right
// indent`, with the FIRST line further inset by `<w:ind w:firstLine>`. Canonical
// table acquisition must use that same retained paragraph geometry when sizing
// the row; otherwise a first-line indent that forces a wrap can size the row for
// FEWER lines than the retained paint tree contains and overflow the next row.
//
// sample-11.docx exposes this: the "City or Town" header cell and the
// "Cedar University" data cells carry `w:firstLine=432` (21.6 pt). At their
// column width the text wraps to two lines when the indent is honored, but the
// (pre-fix) measurer ignored the indent, sized the row for one line, and the
// second line ("Town" / "University") bled into the row below.
//
// This is a DIFFERENTIAL test: two cells identical except for the first-line
// indent. The recording ctx measures every glyph at `fontSize` px, so the
// wrap is deterministic. The indented cell MUST be taller (it wraps to two
// lines); the un-indented cell stays one line. Pre-fix both measured one line
// (equal heights) — so this fails Red and passes Green.

function makeMeasureContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() {
      return font;
    },
    set font(v: string) {
      font = v;
    },
    letterSpacing: '0px',
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
    save() {},
    restore() {},
    measureText2() {},
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

const SECTION: SectionProps = {
  pageWidth: 500,
  pageHeight: 1000,
  marginTop: 0,
  marginRight: 0,
  marginBottom: 0,
  marginLeft: 0,
  headerDistance: 0,
  footerDistance: 0,
  titlePage: false,
  evenAndOddHeaders: false,
};

function para(text: string, indentFirst: number): DocParagraph {
  return {
    alignment: 'left',
    indentLeft: 0,
    indentRight: 0,
    indentFirst,
    spaceBefore: 0,
    spaceAfter: 0,
    lineSpacing: null,
    numbering: null,
    tabStops: [],
    runs: [
      {
        type: 'text',
        text,
        bold: false,
        italic: false,
        underline: false,
        strikethrough: false,
        fontSize: 10,
        color: null,
        fontFamily: 'Times New Roman',
        fontFamilyEastAsia: 'Times New Roman',
        isLink: false,
        background: null,
        vertAlign: null,
        hyperlink: null,
      },
    ],
    defaultFontSize: 10,
    defaultFontFamily: 'Times New Roman',
    widowControl: false,
  } as unknown as DocParagraph;
}

function cellWith(p: DocParagraph): DocTableCell {
  return {
    content: [{ type: 'paragraph', ...p }],
    colSpan: 1,
    vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background: null,
    vAlign: 'top',
    widthPt: 120,
    marginTop: 0,
    marginBottom: 0,
    marginLeft: 0,
    marginRight: 0,
  } as unknown as DocTableCell;
}

function rowWith(p: DocParagraph): DocTableRow {
  return {
    cells: [cellWith(p)],
    rowHeight: null,
    rowHeightRule: 'auto',
    isHeader: false,
  } as unknown as DocTableRow;
}

function table(): DocTable {
  return {
    colWidths: [120],
    rows: [],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 0,
    cellMarginRight: 0,
    jc: 'left',
  } as unknown as DocTable;
}

describe('canonical retained row height — paragraph indent affects wrap (§17.3.1.12 / §17.4.80)', () => {
  const t = table();
  const retainedAdvance = (p: DocParagraph) => {
    const advance = layoutBodyTableRowAdvances(
      { ...t, rows: [rowWith(p)] },
      SECTION,
      makeMeasureContext(),
    )[0];
    if (advance === undefined) throw new Error('Canonical table omitted the row');
    return advance;
  };
  // "WORD ALPHA" = 10 glyphs × 10 px = 100 px. Fits the 120 px column on one
  // line with no indent. With a 60 px first-line indent the first line only has
  // 60 px, so it breaks after "WORD" (40 px) and "ALPHA" wraps to a second line.
  const TEXT = 'WORD ALPHA';

  it('un-indented cell stays one line', () => {
    const h = retainedAdvance(para(TEXT, 0));
    // One 10 px line box (+ its descent). Comfortably below two lines.
    expect(h).toBeLessThan(16);
  });

  it('first-line-indent cell wraps to two lines and the row grows to fit', () => {
    const h1 = retainedAdvance(para(TEXT, 0));
    const h2 = retainedAdvance(para(TEXT, 60));
    // The indented cell wraps to a second line, so its row must be ~2× taller.
    expect(h2).toBeGreaterThan(h1 * 1.8);
  });
});

describe('vertical merge interior border footprint (WORD_VERTICAL_MERGE_INTERIOR_BORDER_FOOTPRINT)', () => {
  const single = { width: 0.5, color: null, style: 'single' };
  const mergedCell = (vMerge: boolean, bottom: typeof single | null) => ({
    ...cellWith(para('A', 0)),
    vMerge,
    borders: { top: null, bottom, left: null, right: null, insideH: null, insideV: null },
  }) as unknown as DocTableCell;
  const plainRow = () => rowWith(para('A', 0));
  const advances = (restartBottom: typeof single | null, continuationBottom: typeof single | null) =>
    layoutBodyTableRowAdvances({
      ...table(),
      rows: [plainRow(),
        { ...plainRow(), cells: [mergedCell(true, restartBottom)] },
        { ...plainRow(), cells: [mergedCell(false, continuationBottom)] },
        plainRow()],
    } as DocTable, SECTION, makeMeasureContext());

  it('charges the first merged cell\'s own bottom border to the next row although no rule is drawn', () => {
    const none = advances(null, null);
    const restart = advances(single, null);
    expect(restart[1]).toBeCloseTo(none[1]!, 10);
    expect(restart[2]).toBeCloseTo(none[2]! + 0.5, 10);
    expect(restart[3]).toBeCloseTo(none[3]!, 10);
  });

  it('charges nothing inside the merge for a border on the last constituent only', () => {
    const none = advances(null, null);
    const terminal = advances(null, single);
    expect(terminal[2]).toBeCloseTo(none[2]!, 10);
    // The terminal border is the rule between the merge and the next row.
    expect(terminal[3]).toBeCloseTo(none[3]! + 0.5, 10);
  });
});
