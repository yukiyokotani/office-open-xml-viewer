import { describe, it, expect } from 'vitest';
import { bodyFragmentFor, computePages, computeColumns } from './renderer.js';
import type {
  BlockContinuationRange,
  TableCellFragmentLayout,
  TableFragmentLayout,
} from './layout/table-pagination.js';
import type {
  BodyElement, CellElement, DocParagraph, DocTableCell, DocxTextRun, ShapeRun, SectionProps, PaginatedBodyElement, DocTable, DocTableRow,
} from './types';

// Unit tests for computePages pagination behaviour that the renderer-path VRT
// (local-only, private samples) cannot guard in CI. A deterministic stub canvas
// makes line wrapping and line heights predictable: glyph advance = charCount ×
// fontPx, and the font box = 0.8/0.2 em (so a single line is exactly fontPx tall
// with no spacing/grid). CJK characters break between any two glyphs, so a run of
// N of them wraps to ceil(N / charsPerLine) lines.

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

type DocRun = DocParagraph['runs'][number];

function para(
  opts: {
    text?: string;
    fontSize?: number;
    widowControl?: boolean;
    keepLines?: boolean;
    keepNext?: boolean;
    spaceBefore?: number;
    spaceAfter?: number;
    indentFirst?: number;
    markVanish?: boolean;
  } = {},
): BodyElement {
  const fontSize = opts.fontSize ?? 20;
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: opts.indentFirst ?? 0,
    spaceBefore: opts.spaceBefore ?? 0, spaceAfter: opts.spaceAfter ?? 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: opts.text ? [textRun(opts.text, fontSize)] : [],
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics',
    widowControl: opts.widowControl,
    keepLines: opts.keepLines,
    keepNext: opts.keepNext,
    markVanish: opts.markVanish,
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

/** A minimal anchored text-box shape (wps:wsp under wp:anchor). `wrapMode`
 *  defaults to topAndBottom so it reserves a full-width float band. */
function shapeRun(opts: {
  widthPt: number;
  heightPt: number;
  anchorYPt?: number;
  anchorYFromPara?: boolean;
  anchorXFromMargin?: boolean;
  wrapMode?: string | null;
  wrapSide?: string | null;
  distTop?: number;
  distBottom?: number;
  distLeft?: number;
  distRight?: number;
}): DocRun {
  const s: ShapeRun = {
    widthPt: opts.widthPt,
    heightPt: opts.heightPt,
    anchorXPt: 0,
    anchorYPt: opts.anchorYPt ?? 0,
    anchorXFromMargin: opts.anchorXFromMargin ?? true,
    anchorYFromPara: opts.anchorYFromPara ?? true,
    zOrder: 0,
    subpaths: [],
    presetGeometry: 'rect',
    fill: { fillType: 'solid', color: 'FFFFFF' },
    stroke: null,
    wrapMode: opts.wrapMode === undefined ? 'topAndBottom' : opts.wrapMode,
    wrapSide: opts.wrapSide ?? null,
    distTop: opts.distTop ?? 0,
    distBottom: opts.distBottom ?? 0,
    distLeft: opts.distLeft ?? 0,
    distRight: opts.distRight ?? 0,
  };
  return { type: 'shape', ...s } as DocRun;
}

/** A paragraph carrying an explicit run list (e.g. an anchored shape, optionally
 *  followed by inline text). */
function paraWith(runs: DocRun[], opts: { fontSize?: number } = {}): BodyElement {
  const fontSize = opts.fontSize ?? 20;
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs,
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics',
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function inlineImageRun(widthPt: number, heightPt: number): DocRun {
  return {
    type: 'image',
    imagePath: 'word/media/image.png',
    mimeType: 'image/png',
    widthPt,
    heightPt,
    anchor: false,
  } as DocRun;
}

/** A single-column block table whose rows each have a fixed `exact` height (pt),
 *  so its measured height is deterministic and independent of cell-content
 *  measurement (resolveTableRowHeights short-circuits on rowHeightRule="exact"). */
function fixedTable(rowHeightsPt: number[]): BodyElement {
  const rows: DocTableRow[] = rowHeightsPt.map((hPt) => ({
    cells: [
      {
        content: [],
        colSpan: 1,
        vMerge: null,
        borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
        background: null,
        vAlign: 'top',
        widthPt: null,
      },
    ],
    rowHeight: hPt,
    rowHeightRule: 'exact',
    isHeader: false,
  }));
  const t: DocTable = {
    colWidths: [70],
    rows,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
  };
  return { type: 'table', ...t } as BodyElement;
}

function wordIgnoredPositionedTable(rowHeightsPt: number[]): BodyElement {
  const table = fixedTable(rowHeightsPt) as BodyElement & {
    tblpPr: NonNullable<DocTable['tblpPr']>;
    __tableLayout: unknown;
  };
  table.tblpPr = {
    leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
    horzAnchor: 'text', horzSpecified: true, vertAnchor: 'margin',
    tblpX: 0, tblpY: 0,
  };
  table.__tableLayout = {
    effectiveStyleId: 'TableNormal',
    ordinaryFlow: true,
    grid: { authored: true, columns: [{ width: '1400' }], requiredColumnCount: 1 },
    preferredWidth: null,
    layout: null,
    cellSpacing: null,
  };
  return table;
}

function autoTableWithTallRow(blockCount: number, rowOverrides: Partial<DocTableRow> = {}): BodyElement {
  const content = Array.from({ length: blockCount }, (_, i) =>
    para({ text: `row ${i}`, fontSize: 20 }) as CellElement,
  );
  const row: DocTableRow = {
    cells: [
      {
        content,
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
    ...rowOverrides,
  };
  const t: DocTable = {
    colWidths: [160],
    rows: [row],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
  };
  return { type: 'table', ...t } as BodyElement;
}

function autoTableWithSingleWrappedParagraph(charCount: number, rowOverrides: Partial<DocTableRow> = {}): BodyElement {
  const row: DocTableRow = {
    cells: [
      {
        content: [para({ text: 'あ'.repeat(charCount), fontSize: 20 }) as CellElement],
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
    ...rowOverrides,
  };
  const t: DocTable = {
    colWidths: [160],
    rows: [row],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
  };
  return { type: 'table', ...t } as BodyElement;
}

function autoTableWithIntroRowThenSingleWrappedParagraph(charCount: number): BodyElement {
  const first = (fixedTable([40]) as unknown as DocTable).rows[0];
  const second = (autoTableWithSingleWrappedParagraph(charCount) as unknown as DocTable).rows[0];
  const t: DocTable = {
    colWidths: [160],
    rows: [first, second],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
  };
  return { type: 'table', ...t } as BodyElement;
}

function autoTableWithIntroRowThenSpacedWrappedParagraph(): BodyElement {
  const first = (fixedTable([45]) as unknown as DocTable).rows[0];
  const second: DocTableRow = {
    cells: [
      {
        content: [
          para({ text: 'first', fontSize: 20, spaceAfter: 10 }) as CellElement,
          para({ text: 'あ'.repeat(32), fontSize: 20, spaceBefore: 8 }) as CellElement,
        ],
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
  };
  const t: DocTable = {
    colWidths: [160],
    rows: [first, second],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
  };
  return { type: 'table', ...t } as BodyElement;
}

function autoTableWithNestedFixedTable(rowHeightsPt: number[]): BodyElement {
  const nested = fixedTable(rowHeightsPt) as unknown as DocTable;
  const row: DocTableRow = {
    cells: [
      {
        content: [{ type: 'table', ...nested } as unknown as CellElement],
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
  };
  const t: DocTable = {
    colWidths: [160],
    rows: [row],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
  };
  return { type: 'table', ...t } as BodyElement;
}

const sliceOf = (el: PaginatedBodyElement) =>
  (el as { lineSlice?: { start: number; end: number } }).lineSlice;

const colOf = (el: PaginatedBodyElement) => el.colIndex;

/** A body-level column break (ECMA-376 §17.3.1.20, hoisted by the parser). */
const colBreak = (): BodyElement => ({ type: 'columnBreak' } as BodyElement);
const pageBreak = (): BodyElement => ({ type: 'pageBreak' } as BodyElement);

/** A body-level section break (ECMA-376 §17.6.x). `columns` is the ENDING
 *  section's `<w:cols>` (null/undefined ⇒ single full-width column). */
const sectionBreak = (
  kind: 'continuous' | 'nextPage' | 'oddPage' | 'evenPage',
  columns: SectionProps['columns'] = null,
): BodyElement => ({ type: 'sectionBreak', kind, columns } as BodyElement);

/** Text of a paragraph element (joins its text runs). */
const textOf = (el: PaginatedBodyElement): string =>
  el.type === 'paragraph'
    ? (el as unknown as DocParagraph).runs
        .filter((r) => r.type === 'text')
        .map((r) => (r as DocxTextRun).text)
        .join('')
    : '';

const paragraphOccurrenceWithRun = (
  page: PaginatedBodyElement[],
  run: DocRun,
): PaginatedBodyElement | undefined => page.find(
  (element) => element.type === 'paragraph'
    && (element as unknown as DocParagraph).runs.includes(run),
);

function retainedTableOn(page: PaginatedBodyElement[]) {
  const element = page.find((el) => el.type === 'table');
  if (!element) return undefined;
  const placed = bodyFragmentFor(element);
  if (placed?.fragment.kind !== 'table' || !('flowBounds' in placed.fragment)) {
    throw new Error('expected a retained table fragment');
  }
  return { element, placed, fragment: placed.fragment as TableFragmentLayout };
}

function bodyPaginationGeometry(pages: PaginatedBodyElement[][]) {
  return {
    pageCount: pages.length,
    pages: pages.map((page) => page.map((el) => {
      const placed = bodyFragmentFor(el);
      const paragraph = placed?.fragment.kind === 'paragraph' ? placed.fragment : undefined;
      return {
        type: el.type,
        lineSlice: sliceOf(el) ?? null,
        columnIndex: el.colIndex ?? null,
        columnTopPt: el.colTopPt ?? null,
        measuredLineTopsPt: paragraph?.lines.map((line) => line.bounds.yPt ?? null) ?? null,
        measuredWithFloats: paragraph ? paragraph.exclusions.length > 0 : null,
      };
    })),
  };
}

describe('computePages — shared paragraph measurement migration geometry', () => {
  it('preserves body page count, line ranges, and measured float top placement', () => {
    const longText = 'あ'.repeat(72);
    const followingText = '終'.repeat(16);
    const body = [
      paraWith([
        shapeRun({ widthPt: 160, heightPt: 35, anchorYPt: 0, wrapMode: 'topAndBottom' }),
      ]),
      para({
        text: longText,
        fontSize: 20,
        widowControl: false,
        spaceBefore: 5,
        spaceAfter: 7,
      }),
      para({ text: followingText, fontSize: 20 }),
    ];
    const pages = computePages(body, section(), makeCtx());
    const geometry = bodyPaginationGeometry(pages);

    expect(geometry.pageCount).toBe(3);
    expect(geometry.pages.map((page) => page.map((el) => el.lineSlice))).toEqual([
      [null, { start: 0, end: 2 }],
      [{ start: 2, end: 7 }],
      [{ start: 7, end: 9 }, null],
    ]);
    expect(geometry.pages.flatMap((page) => page)
      .filter((el) => el.lineSlice !== null)
      .map((el) => el.measuredLineTopsPt)).toEqual([
      [80, 100],
      // A continuation owns page-local retained geometry. Carrying the source
      // paragraph's absolute Y values into later pages made its ink overlap the
      // next story even though the paginator advanced by the correct slice.
      [75, 95, 115, 135, 155],
      [75, 95],
    ]);
    expect(pages.findIndex((page) => page.some((el) => textOf(el) === followingText))).toBe(2);
  });

  it('remeasures destination placement when the first line cannot fit beside a page float', () => {
    const targetText = 'あ'.repeat(16);
    const body = [
      paraWith([
        shapeRun({ widthPt: 160, heightPt: 35, anchorYPt: 0, wrapMode: 'topAndBottom' }),
      ]),
      para({
        text: targetText,
        fontSize: 20,
        widowControl: false,
        spaceBefore: 30,
      }),
    ];

    const pages = computePages(body, section(), makeCtx());
    const geometry = bodyPaginationGeometry(pages);
    const target = geometry.pages[1]?.find((el) => el.lineSlice?.start === 0);

    expect(geometry.pageCount).toBe(2);
    expect(target?.lineSlice).toEqual({ start: 0, end: 2 });
    expect(target?.measuredWithFloats).toBe(false);
    expect(target?.measuredLineTopsPt).toEqual([50, 70]);
  });

  it('remeasures an unplaced paragraph at the next unequal-width column', () => {
    const targetText = 'あ'.repeat(4);
    const unequalColumns = section({
      columns: {
        count: 2,
        spacePt: 0,
        equalWidth: false,
        sep: false,
        cols: [
          { widthPt: 100, spacePt: 12 },
          { widthPt: 48, spacePt: 0 },
        ],
      },
    });
    const body = [
      ...Array.from({ length: 4 }, () => para({ text: 'a', fontSize: 20 })),
      para({ text: 'b', fontSize: 10 }),
      para({ text: targetText, fontSize: 20, widowControl: false }),
    ];

    const pages = computePages(body, unequalColumns, makeCtx());
    const target = pages[0]?.find((el) => textOf(el) === targetText);
    const placed = target ? bodyFragmentFor(target) : undefined;
    const retained = placed?.fragment.kind === 'paragraph' ? placed.fragment : undefined;

    expect(pages).toHaveLength(1);
    expect(target?.colIndex).toBe(1);
    expect(target?.colTopPt).toBe(20);
    expect(target?.lineSlice).toEqual({ start: 0, end: 2 });
    expect(placed?.widthPt).toBe(48);
    expect(retained?.lines.map((line) => line.bounds.yPt ?? null)).toEqual([20, 40]);
  });

  // ECMA-376 §17.6.4 (cols / equalWidth=false / per-<w:col> w:w) — a newspaper
  // column is a text region of a DEFINED width; text must wrap to ITS column's
  // width. §17.3.1.12 (ind) — first-line/hanging indent applies to the paragraph's
  // FIRST line only, so a continuation into a later column must not re-trigger it.
  //
  // Defect: `splitParagraphAcrossPages` measures the paragraph ONCE at its FIRST
  // placement (the WIDE column) and slices that single partition across columns.
  // `remeasureBeforeFirstLine` only re-measures when NOTHING has been placed yet
  // (`lineIdx === 0`); the ordinary continuation branch (`if (!isFinalSlice)
  // { newPage(...) }`, and the `lastFitting === firstFitting` advance for
  // `lineIdx > 0`) advances into the NARROW column WITHOUT re-wrapping, so the
  // remainder keeps the wide column's line breaks and overflows the narrow column.
  //
  // Fixed by the issue-#908 boundary-provenance seam: `layoutLines` records each
  // line's consumed-content END boundary (break-aware, in field-resolved segment-
  // stream coordinates), `measureParagraph` lays out a remainder suffix from such a
  // boundary (first-line indent suppressed per §17.3.1.12 — a continuation is not a
  // first line), and `splitParagraphAcrossPages` re-measures the remainder when the
  // destination placement mismatches the recorded one (same-width continuations keep
  // the exact single-measurement path).
  it('re-wraps a paragraph continuation into a narrower unequal-width column', () => {
    // col0 = 100pt wide (5 CJK glyphs/line at 20pt), col1 = 48pt wide (2/line).
    // 100 + 12 + 48 = 160 = content width; both bands are 100pt tall (5 lines).
    const unequalColumns = section({
      columns: {
        count: 2,
        spacePt: 0,
        equalWidth: false,
        sep: false,
        cols: [
          { widthPt: 100, spacePt: 12 },
          { widthPt: 48, spacePt: 0 },
        ],
      },
    });
    // 35 glyphs: 7 lines at the WIDE width. col0 holds 5 (chars 0–24); the tail
    // (chars 25–34, 10 glyphs) continues into col1 and MUST re-wrap to 2/line.
    const body = [para({ text: 'あ'.repeat(35), fontSize: 20, widowControl: false })];

    const pages = computePages(body, unequalColumns, makeCtx());
    expect(pages).toHaveLength(1);

    // Every slice of the (single) paragraph, with its retained line geometry.
    const slices = paragraphSlices(pages);

    const col0Lines = slices.filter((s) => s.colIndex === 0).flatMap(paintedLines);
    const col1Lines = slices.filter((s) => s.colIndex === 1).flatMap(paintedLines);

    // The paragraph genuinely spans both columns (precondition, not the assertion
    // under test): col0 filled its 5 lines, and there IS continuation in col1.
    expect(col0Lines.length).toBe(5);
    expect(col1Lines.length).toBeGreaterThan(0);

    // §17.6.4 — NO continuation line may overflow col1's 48pt band.
    const EPS = 1e-6;
    for (const line of col1Lines) {
      expect(lineWidth(line)).toBeLessThanOrEqual(48 + EPS);
    }
    // The remainder actually re-wrapped: col1 packs at most 2 glyphs/line (its
    // width), strictly narrower than col0's 5/line. With the defect the tail keeps
    // col0's 5-glyph, 100pt-wide lines (which overflow the 48pt column).
    expect(Math.max(...col1Lines.map(lineChars))).toBeLessThanOrEqual(2);
    expect(Math.max(...col1Lines.map(lineChars))).toBeLessThan(
      Math.max(...col0Lines.map(lineChars)),
    );
    // All 35 glyphs are still placed (10 in col1), and the continuation records the
    // NARROW column width (no stale wide-column measurement leaks through).
    expect(col1Lines.reduce((sum, l) => sum + lineChars(l), 0)).toBe(10);
    for (const s of slices.filter((sl) => sl.colIndex === 1)) {
      expect(bodyFragmentFor(s)?.widthPt).toBe(48);
    }
  });

  it('keeps state-sensitive continuations on the original partition', () => {
    const unequalColumns = section({
      columns: {
        count: 2,
        spacePt: 0,
        equalWidth: false,
        sep: false,
        cols: [
          { widthPt: 100, spacePt: 12 },
          { widthPt: 48, spacePt: 0 },
        ],
      },
    });
    const fieldPara = para({
      text: 'あ'.repeat(35),
      fontSize: 20,
      widowControl: false,
    }) as unknown as DocParagraph;
    (fieldPara.runs as unknown[]).push({
      type: 'field', fieldType: 'numPages', instruction: 'NUMPAGES', fallbackText: '?',
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', background: null,
    });

    const slices = paragraphSlices(computePages(
      [fieldPara as unknown as BodyElement],
      unequalColumns,
      makeCtx(),
    ));
    const col1Slices = slices.filter((slice) => slice.colIndex === 1);

    expect(col1Slices.length).toBeGreaterThan(0);
    expect.soft(slices.every((slice) => slice.lineSlice?.continues === undefined)).toBe(true);
    expect.soft(col1Slices.every((slice) => (slice.lineSlice?.start ?? 0) > 0)).toBe(true);
  });

  it('keeps numbered continuations on the original partition', () => {
    const unequalColumns = section({
      columns: {
        count: 2,
        spacePt: 0,
        equalWidth: false,
        sep: false,
        cols: [
          { widthPt: 100, spacePt: 12 },
          { widthPt: 48, spacePt: 0 },
        ],
      },
    });
    const numberedPara = para({
      text: 'あ'.repeat(35),
      fontSize: 20,
      widowControl: false,
    }) as unknown as DocParagraph;
    numberedPara.numbering = {
      numId: 1, level: 0, format: 'decimal', text: '1.',
      indentLeft: 0, tab: 18, suff: 'tab', jc: 'left',
    } as unknown as DocParagraph['numbering'];

    const slices = paragraphSlices(computePages(
      [numberedPara as unknown as BodyElement],
      unequalColumns,
      makeCtx(),
    ));
    const col1Slices = slices.filter((slice) => slice.colIndex === 1);

    expect(col1Slices.length).toBeGreaterThan(0);
    expect.soft(slices.every((slice) => slice.lineSlice?.continues === undefined)).toBe(true);
    expect.soft(col1Slices.every((slice) => (slice.lineSlice?.start ?? 0) > 0)).toBe(true);
  });

  type RuntimeLine = NonNullable<ReturnType<typeof retainedParagraphOf>>['lines'][number];
  type RuntimeSlice = PaginatedBodyElement;
  const retainedParagraphOf = (slice: RuntimeSlice) => {
    const placed = bodyFragmentFor(slice);
    return placed?.fragment.kind === 'paragraph' ? placed.fragment : undefined;
  };
  const paragraphSlices = (pages: PaginatedBodyElement[][]): RuntimeSlice[] =>
    pages.flat().filter((el) => el.type === 'paragraph') as RuntimeSlice[];
  const paintedLines = (slice: RuntimeSlice): RuntimeLine[] => {
    const paragraph = retainedParagraphOf(slice);
    return [...(paragraph?.lines ?? [])];
  };
  const lineWidth = (line: RuntimeLine): number =>
    line.placements.reduce(
      (sum, placement) => sum + ('advancePt' in placement ? placement.advancePt : 0),
      0,
    );
  const lineChars = (line: RuntimeLine): number =>
    line.placements.reduce(
      (sum, placement) => sum + (placement.kind === 'text' ? [...placement.text].length : 0),
      0,
    );

  it('keeps equal-width column continuations on the original line partition', () => {
    const equalColumns = section({
      columns: {
        count: 2,
        spacePt: 20,
        equalWidth: true,
        sep: false,
        cols: [],
      },
    });
    const sharedWidth = computeColumns(equalColumns)[0].wPt;
    const pages = computePages(
      [para({ text: 'あ'.repeat(35), fontSize: 20, widowControl: false })],
      equalColumns,
      makeCtx(),
    );
    const slices = paragraphSlices(pages);
    const col1Slices = slices.filter((slice) => slice.colIndex === 1);

    expect(col1Slices.length).toBeGreaterThan(0);
    expect(col1Slices.every((slice) => (slice.lineSlice?.start ?? 0) > 0)).toBe(true);
    expect(slices.every((slice) => slice.lineSlice?.continues === undefined)).toBe(true);
    expect(slices.every((slice) => {
      const range = slice.lineSlice;
      return range === undefined || paintedLines(slice).length === range.end - range.start;
    })).toBe(true);
    expect(slices.every((slice) => bodyFragmentFor(slice)?.widthPt === sharedWidth)).toBe(true);
  });

  it('composes remainder boundaries across three unequal-width columns', () => {
    // 100 + 6 + 48 + 6 + 30 = 190pt, exactly the section content width.
    const unequalColumns = section({
      pageWidth: 230,
      columns: {
        count: 3,
        spacePt: 0,
        equalWidth: false,
        sep: false,
        cols: [
          { widthPt: 100, spacePt: 6 },
          { widthPt: 48, spacePt: 6 },
          { widthPt: 30, spacePt: 0 },
        ],
      },
    });
    const pages = computePages(
      [para({ text: 'あ'.repeat(40), fontSize: 20, widowControl: false })],
      unequalColumns,
      makeCtx(),
    );
    const slices = paragraphSlices(pages);
    const col1Lines = slices.filter((slice) => slice.colIndex === 1).flatMap(paintedLines);
    const col2Slices = slices.filter((slice) => slice.colIndex === 2);
    const col2Lines = col2Slices.flatMap(paintedLines);

    expect(col1Lines.length).toBeGreaterThan(0);
    expect(col2Lines.length).toBeGreaterThan(0);
    expect(col1Lines.every((line) => lineWidth(line) <= 48 + 1e-6)).toBe(true);
    expect(col2Lines.every((line) => lineWidth(line) <= 30 + 1e-6)).toBe(true);
    expect(slices.flatMap(paintedLines).reduce((sum, line) => sum + lineChars(line), 0)).toBe(40);
    expect(col2Slices.every((slice) => bodyFragmentFor(slice)?.widthPt === 30)).toBe(true);
  });

  it('does not re-apply first-line indent when re-wrapping a continuation', () => {
    const unequalColumns = section({
      columns: {
        count: 2,
        spacePt: 0,
        equalWidth: false,
        sep: false,
        cols: [
          { widthPt: 100, spacePt: 12 },
          { widthPt: 48, spacePt: 0 },
        ],
      },
    });
    const pages = computePages(
      [para({
        text: 'あ'.repeat(35),
        fontSize: 20,
        widowControl: false,
        indentFirst: 20,
      })],
      unequalColumns,
      makeCtx(),
    );
    const col1Lines = paragraphSlices(pages)
      .filter((slice) => slice.colIndex === 1)
      .flatMap(paintedLines);
    const glyphCounts = col1Lines.map(lineChars);

    expect(glyphCounts.length).toBeGreaterThan(1);
    expect(glyphCounts.every((count) => count === glyphCounts[0])).toBe(true);
    expect(col1Lines.every((line) => lineWidth(line) <= 48 + 1e-6)).toBe(true);
  });

  it('does not re-charge leading spacing on a re-wrapped continuation', () => {
    const unequalColumns = section({
      columns: {
        count: 2,
        spacePt: 0,
        equalWidth: false,
        sep: false,
        cols: [
          { widthPt: 100, spacePt: 12 },
          { widthPt: 48, spacePt: 0 },
        ],
      },
    });
    const pages = computePages(
      [para({
        text: 'あ'.repeat(35),
        fontSize: 20,
        widowControl: false,
        spaceBefore: 30,
      })],
      unequalColumns,
      makeCtx(),
    );
    const firstCol1Slice = paragraphSlices(pages).find((slice) => slice.colIndex === 1);
    const placed = firstCol1Slice ? bodyFragmentFor(firstCol1Slice) : undefined;

    expect(firstCol1Slice?.lineSlice?.continues).toBe(true);
    expect(placed?.fragment.kind).toBe('paragraph');
    if (placed?.fragment.kind === 'paragraph') {
      expect(placed.fragment.spacing.beforePt).toBe(0);
    }
  });
});

describe('computePages — empty-paragraph relocation (C2: §17.3.1.29)', () => {
  it('moves an unsplittable mark-only paragraph to the next page instead of overflowing the bottom margin', () => {
    // content height = 140 - 40 = 100; each empty mark = 20px → exactly 5 per page.
    const body = Array.from({ length: 7 }, () => para()); // 7 empty paragraphs
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(2);
    expect(pages[0].length).toBe(5); // page 1 fills exactly
    expect(pages[1].length).toBe(2); // overflow relocated, NOT clipped onto page 1
    // no page holds more than its 5-line capacity (would mean an overflow)
    for (const p of pages) expect(p.length).toBeLessThanOrEqual(5);
  });
});

describe('computePages — tall header/footer reserve indexing (§17.6.11)', () => {
  it('applies a uniform reserve to EVERY page, not just those in the array (pass-2 page growth)', () => {
    // The two-pass paginator computes reserves from pass 1, then re-paginates; a tall
    // header shrinks every page from the top, so pass 2 routinely has MORE pages than
    // pass 1. A page index past the reserve array must clamp to the last entry (the
    // uniform reserve), not fall to zero — else later pages stay un-reserved, over-pack,
    // and the down-shifted body overflows the bottom margin.
    // content height = 100; empty mark = 20px → 5/page bare, 4/page with a 20px top reserve.
    const body = Array.from({ length: 14 }, () => para());
    const reserve = { top: 20, bottom: 0 };
    // A single-entry array expresses "uniform 20px top reserve" (one page measured).
    const short = computePages(body, section(), makeCtx(), {}, undefined, [], [reserve]);
    // The same reserve spelled out for more pages than the body can ever need.
    const full = computePages(body, section(), makeCtx(), {}, undefined, [], Array.from({ length: 14 }, () => reserve));
    // The short (clamped) array must paginate identically to the fully-specified one.
    expect(short.map((p) => p.length)).toEqual(full.map((p) => p.length));
    // Non-triviality: the reserve really is in force (4 per page, not the bare 5) and
    // the body genuinely spans several pages where the clamp matters.
    expect(short[0].length).toBe(4);
    expect(short.length).toBeGreaterThanOrEqual(4);
  });
});

describe('computePages — Word-effective table flow ([MS-OI29500] 2.1.162(b-c))', () => {
  it('charges an authored but ignored tblpPr as ordinary block-table flow', () => {
    const pages = computePages([
      wordIgnoredPositionedTable([40]),
      para({ text: 'after', fontSize: 20 }),
    ], section(), makeCtx());
    const tableOccurrence = pages[0]?.find((element) => element.type === 'table');
    const paragraphOccurrence = pages[0]?.find((element) => element.type === 'paragraph');
    const tablePlacement = tableOccurrence ? bodyFragmentFor(tableOccurrence) : undefined;
    const paragraphPlacement = paragraphOccurrence ? bodyFragmentFor(paragraphOccurrence) : undefined;

    expect(tablePlacement).toBeDefined();
    expect(paragraphPlacement?.yPt).toBeGreaterThanOrEqual(
      (tablePlacement?.yPt ?? 0) + (tablePlacement?.heightPt ?? 0),
    );
  });
});

describe('computePages — over-tall table row splitting (§17.4 table pagination)', () => {
  const paragraphContinuation = (cell: TableCellFragmentLayout, blockIndex = 0) => {
    const block = cell.blocks.find((candidate) => (
      candidate.layout.kind === 'paragraph'
      && cell.contentRanges.some((range) => (
        range.kind === 'paragraph' && range.blockIndex === blockIndex
      ))
    ));
    return block?.layout.kind === 'paragraph' ? block.layout.continuation : undefined;
  };

  it('splits an auto-height row by cell block boundaries instead of overflowing the footer band', () => {
    const pages = computePages([autoTableWithTallRow(8)], section(), makeCtx());
    const tables = pages.map(retainedTableOn).filter((table) => table !== undefined);

    expect(tables.length).toBeGreaterThan(1);
    expect(tables.flatMap((table) => table.fragment.rows)
      .map((row) => [row.logicalRowIndex, row.fragmentIndex])).toEqual([[0, 0], [0, 1]]);
    expect(tables.flatMap((table) => table.fragment.rows[0]?.cells[0]?.contentRanges ?? []))
      .toEqual(Array.from({ length: 8 }, (_, blockIndex) => ({
        kind: 'paragraph', blockIndex, lineStart: 0, lineEnd: 1,
      })));
    for (const table of tables) {
      expect(table.placed.heightPt).toBeLessThanOrEqual(100);
    }
  });

  it('splits a splittable auto-height row into the remaining page band before continuing', () => {
    const pages = computePages([para(), para(), autoTableWithTallRow(4)], section(), makeCtx());
    const firstPageTable = retainedTableOn(pages[0]);
    const secondPageTable = retainedTableOn(pages[1]);

    expect(firstPageTable).toBeDefined();
    expect(secondPageTable).toBeDefined();

    const firstSliceBlocks = firstPageTable?.fragment.rows[0]?.cells[0]?.contentRanges.length ?? 0;
    const secondSliceBlocks = secondPageTable?.fragment.rows[0]?.cells[0]?.contentRanges.length ?? 0;
    expect(firstSliceBlocks).toBeGreaterThan(0);
    expect(firstSliceBlocks).toBeLessThan(4);
    expect(firstSliceBlocks + secondSliceBlocks).toBe(4);
    expect([
      firstPageTable?.fragment.rows[0]?.fragmentIndex,
      secondPageTable?.fragment.rows[0]?.fragmentIndex,
    ]).toEqual([0, 1]);
  });

  it('keeps a cantSplit row intact when it does not fit the remaining page band', () => {
    const pages = computePages([para(), para(), autoTableWithTallRow(4, { cantSplit: true })], section(), makeCtx());

    expect(pages[0].some((el) => el.type === 'table')).toBe(false);
    const secondPageTable = retainedTableOn(pages[1]);
    expect(secondPageTable?.fragment.rows[0]?.cells[0]?.contentRanges).toEqual(
      Array.from({ length: 4 }, (_, blockIndex) => ({ kind: 'whole', blockIndex })),
    );
  });

  it('splits a single cell paragraph in a splittable row at line boundaries', () => {
    const pages = computePages([para(), para(), autoTableWithSingleWrappedParagraph(32)], section(), makeCtx());
    const firstPageTable = retainedTableOn(pages[0]);
    const secondPageTable = retainedTableOn(pages[1]);

    expect(firstPageTable).toBeDefined();
    expect(secondPageTable).toBeDefined();

    const firstCell = firstPageTable?.fragment.rows[0]?.cells[0];
    const secondCell = secondPageTable?.fragment.rows[0]?.cells[0];
    expect(firstCell?.contentRanges).toEqual([
      { kind: 'paragraph', blockIndex: 0, lineStart: 0, lineEnd: 3 },
    ]);
    expect(secondCell?.contentRanges).toEqual([
      { kind: 'paragraph', blockIndex: 0, lineStart: 3, lineEnd: 4 },
    ]);
    expect(firstCell && paragraphContinuation(firstCell)).toEqual({
      lineStart: 0, lineEnd: 3, continuesFromPrevious: false, continuesOnNext: true,
    });
    expect(secondCell && paragraphContinuation(secondCell)).toEqual({
      lineStart: 3, lineEnd: 4, continuesFromPrevious: true, continuesOnNext: false,
    });
  });

  it('splits the next row when it overflows after earlier rows filled part of the page', () => {
    const pages = computePages([autoTableWithIntroRowThenSingleWrappedParagraph(32)], section(), makeCtx());
    const firstPageTable = retainedTableOn(pages[0]);
    const secondPageTable = retainedTableOn(pages[1]);

    expect(firstPageTable?.fragment.rows.map((row) => [row.logicalRowIndex, row.fragmentIndex]))
      .toEqual([[0, 0], [1, 0]]);
    expect(secondPageTable?.fragment.rows.map((row) => [row.logicalRowIndex, row.fragmentIndex]))
      .toEqual([[1, 1]]);

    expect(firstPageTable?.fragment.rows[1]?.cells[0]?.contentRanges).toEqual([
      { kind: 'paragraph', blockIndex: 0, lineStart: 0, lineEnd: 3 },
    ]);
    expect(secondPageTable?.fragment.rows[0]?.cells[0]?.contentRanges).toEqual([
      { kind: 'paragraph', blockIndex: 0, lineStart: 3, lineEnd: 4 },
    ]);
  });

  it('uses collapsed paragraph spacing when fitting a cell paragraph line into a split row', () => {
    const pages = computePages([autoTableWithIntroRowThenSpacedWrappedParagraph()], section(), makeCtx());
    const firstPageTable = retainedTableOn(pages[0]);
    const continuationTables = pages.slice(1).map(retainedTableOn)
      .filter((table) => table !== undefined);

    expect(firstPageTable?.fragment.rows.map((row) => [row.logicalRowIndex, row.fragmentIndex]))
      .toEqual([[0, 0], [1, 0]]);
    const ranges = [firstPageTable, ...continuationTables]
      .flatMap((table) => table?.fragment.rows ?? [])
      .filter((row) => row.logicalRowIndex === 1)
      .map((row) => row.cells[0]?.contentRanges ?? []);
    expect(ranges).toEqual([
      [
        { kind: 'paragraph', blockIndex: 0, lineStart: 0, lineEnd: 1 },
        { kind: 'paragraph', blockIndex: 1, lineStart: 0, lineEnd: 1 },
      ],
      [{ kind: 'paragraph', blockIndex: 1, lineStart: 1, lineEnd: 4 }],
    ]);
  });

  it('splits a nested table inside a cell at inner row boundaries', () => {
    const pages = computePages([autoTableWithNestedFixedTable([30, 30, 30, 30])], section(), makeCtx());
    expect(pages.length).toBe(2);
    const firstPageTable = retainedTableOn(pages[0]);
    const secondPageTable = retainedTableOn(pages[1]);

    const firstNested = firstPageTable?.fragment.rows[0]?.cells[0]?.blocks[0]?.layout;
    const secondNested = secondPageTable?.fragment.rows[0]?.cells[0]?.blocks[0]?.layout;
    expect(firstPageTable?.fragment.rows[0]?.cells[0]?.contentRanges).toEqual([
      { kind: 'nested-table', blockIndex: 0, childFragmentIndex: 0 },
    ]);
    expect(secondPageTable?.fragment.rows[0]?.cells[0]?.contentRanges).toEqual([
      { kind: 'nested-table', blockIndex: 0, childFragmentIndex: 1 },
    ]);
    expect(firstNested?.kind).toBe('table');
    expect(secondNested?.kind).toBe('table');
    if (firstNested?.kind !== 'table' || secondNested?.kind !== 'table') {
      throw new Error('expected retained nested-table fragments');
    }
    expect(firstNested?.rows.length).toBe(3);
    expect(secondNested?.rows.length).toBe(1);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Mid-page splitting of a row whose cells start a vertical merge (§17.4.84 +
// §17.4.6). The retained fragments preserve each source cell's restart/continue
// semantics. A page-local `visualMergeOwnership` marker supplies only the paint
// ownership needed when a continuation row starts a page; pagination never
// rewrites a continuation into an authored restart.
// ─────────────────────────────────────────────────────────────────────────────
describe('computePages — mid-page split of rows with vMerge restart cells (§17.4.84 + §17.4.6)', () => {
  const emptyBorders = () => ({ top: null, bottom: null, left: null, right: null, insideH: null, insideV: null });
  const mkCell = (content: CellElement[], vMerge: boolean | null, vAlign: 'top' | 'center' | 'bottom' = 'top'): DocTableCell => ({
    content, colSpan: 1, vMerge, borders: emptyBorders(), background: null, vAlign, widthPt: null,
  } as unknown as DocTableCell);
  const shortPara = (text: string) => para({ text, fontSize: 20 }) as CellElement;
  /** [40pt label col | 120pt content col]; row0 = restart label + wrapping
   *  content (24 glyphs at 6/line ⇒ 4 lines = 80pt); rows 1-2 = continue label +
   *  short cell (20pt each). */
  const restartSpanTable = (labelParas: number, labelVAlign: 'top' | 'center' | 'bottom' = 'top'): BodyElement => {
    const label = mkCell(Array.from({ length: labelParas }, (_v, i) => shortPara(`L${i}`)), true, labelVAlign);
    const content = mkCell([para({ text: 'あ'.repeat(24), fontSize: 20 }) as CellElement], null);
    const contRow = (t: string): DocTableRow => ({
      cells: [mkCell([para({}) as CellElement], false), mkCell([shortPara(t)], null)],
      rowHeight: null, rowHeightRule: 'auto', isHeader: false,
    } as unknown as DocTableRow);
    const t: DocTable = {
      colWidths: [40, 120],
      rows: [
        { cells: [label, content], rowHeight: null, rowHeightRule: 'auto', isHeader: false } as unknown as DocTableRow,
        contRow('x'), contRow('y'),
      ],
      borders: emptyBorders(),
      cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
      jc: 'left', layout: 'fixed',
    } as unknown as DocTable;
    return { type: 'table', ...t } as BodyElement;
  };
  const paragraphRanges = (cell: TableCellFragmentLayout | undefined) =>
    (cell?.contentRanges.filter((range): range is Extract<BlockContinuationRange, { kind: 'paragraph' }> => (
      range.kind === 'paragraph'
    )) ?? []);
  const sourceBlockIndices = (cell: TableCellFragmentLayout | undefined) =>
    (cell?.contentRanges.map((range) => range.blockIndex) ?? []);

  it('splits a restart-cell row mid-page without changing its semantic merge role', () => {
    // 2 filler paragraphs (40pt) leave 60pt: the content cell fits 3 of its 4
    // lines; the 2-paragraph label (40pt) fits page 1 entirely.
    const pages = computePages([para(), para(), restartSpanTable(2)], section(), makeCtx());

    const first = retainedTableOn(pages[0]);
    const continuations = pages.slice(1).map(retainedTableOn)
      .filter((table) => table !== undefined);
    const second = continuations[0];
    expect(first).toBeDefined();
    expect(second).toBeDefined();

    expect(first?.fragment.rows.map((row) => [row.logicalRowIndex, row.fragmentIndex]))
      .toEqual([[0, 0]]);
    expect(first?.fragment.rows[0]?.cells[0]?.verticalMerge).toBe('restart');
    expect(sourceBlockIndices(first?.fragment.rows[0]?.cells[0])).toEqual([0, 1]);
    expect(paragraphRanges(first?.fragment.rows[0]?.cells[1])).toEqual([
      { kind: 'paragraph', blockIndex: 0, lineStart: 0, lineEnd: 3 },
    ]);

    const continuationRows = continuations.flatMap((table) => table.fragment.rows);
    expect(continuationRows.map((row) => [row.logicalRowIndex, row.fragmentIndex]))
      .toEqual([[0, 1], [1, 0], [2, 0]]);
    expect(continuationRows.map((row) => row.cells[0]?.verticalMerge)).toEqual([
      'restart', 'continue', 'continue',
    ]);
    expect(second?.fragment.rows[0]?.cells[0]?.contentRanges).toEqual([]);
    expect(paragraphRanges(second?.fragment.rows[0]?.cells[1])).toEqual([
      { kind: 'paragraph', blockIndex: 0, lineStart: 3, lineEnd: 4 },
    ]);
    const firstParagraph = first?.fragment.rows[0]?.cells[1]?.blocks[0]?.layout;
    const secondParagraph = second?.fragment.rows[0]?.cells[1]?.blocks[0]?.layout;
    expect(firstParagraph?.kind === 'paragraph' ? firstParagraph.continuation : undefined).toEqual({
      lineStart: 0, lineEnd: 3, continuesFromPrevious: false, continuesOnNext: true,
    });
    expect(secondParagraph?.kind === 'paragraph' ? secondParagraph.continuation : undefined).toEqual({
      lineStart: 3, lineEnd: 4, continuesFromPrevious: true, continuesOnNext: false,
    });
    const firstSemanticContinuation = continuationRows.find((row) => row.logicalRowIndex === 1);
    expect(firstSemanticContinuation?.cells[0]?.visualMergeOwnership).toBe('continuation');
  });

  it('splits the restart cell content itself when it exceeds the page-1 band', () => {
    // 6 label paragraphs (120pt) cannot fit the 60pt band: 3 land on page 1 and
    // 3 continue in a second fragment with the same authored restart semantics.
    const pages = computePages([para(), para(), restartSpanTable(6)], section(), makeCtx());

    const first = retainedTableOn(pages[0]);
    const second = retainedTableOn(pages[1]);
    expect(sourceBlockIndices(first?.fragment.rows[0]?.cells[0])).toEqual([0, 1, 2]);
    expect(sourceBlockIndices(second?.fragment.rows[0]?.cells[0])).toEqual([3, 4, 5]);
    expect(paragraphRanges(first?.fragment.rows[0]?.cells[1])).toEqual([
      { kind: 'paragraph', blockIndex: 0, lineStart: 0, lineEnd: 3 },
    ]);
    expect(paragraphRanges(second?.fragment.rows[0]?.cells[1])).toEqual([
      { kind: 'paragraph', blockIndex: 0, lineStart: 3, lineEnd: 4 },
    ]);
    for (const pg of pages) {
      const table = retainedTableOn(pg);
      if (table) expect(table.placed.heightPt).toBeLessThanOrEqual(100 + 1e-6);
    }
  });

  it('re-derives the span extension once after the split (no double counting)', () => {
    // ECMA-376 §17.4.84: restart content 200pt, normal cell 80pt (4
    // lines), two 20pt continue rows, 100pt page body. The ORIGINAL heights
    // resolver extended the merge-END row by 200 − (80+20+20) = 80pt; after
    // the split that extension must be RE-DERIVED against the remaining
    // restart content, not left in place while the pieces also carry the
    // fitted content — the double count overflowed the body band and leaked
    // an extra page.
    const pages = computePages([restartSpanTable(10)], section(), makeCtx());
    const tables = pages.map(retainedTableOn).filter((table) => table !== undefined);

    // Every page's table charges at most the 100pt body band.
    for (const table of tables) {
      expect(table.placed.heightPt).toBeLessThanOrEqual(100 + 1e-6);
    }

    // Stable source block/line ranges prove that every label paragraph and
    // content line is owned by exactly one retained fragment.
    const labelBlocks = tables.flatMap((table) => table.fragment.rows)
      .filter((row) => row.logicalRowIndex === 0)
      .flatMap((row) => sourceBlockIndices(row.cells[0]));
    expect(labelBlocks).toEqual(Array.from({ length: 10 }, (_v, index) => index));
    const coveredLines = tables.flatMap((table) => table.fragment.rows)
      .filter((row) => row.logicalRowIndex === 0)
      .flatMap((row) => paragraphRanges(row.cells[1]))
      .flatMap((range) => Array.from(
        { length: range.lineEnd - range.lineStart },
        (_v, offset) => range.lineStart + offset,
      ));
    expect(coveredLines).toEqual([0, 1, 2, 3]);
  });

  it('preserves center vAlign on every retained piece of a split restart cell', () => {
    for (const labelParas of [2, 6]) {
      const pages = computePages([para(), para(), restartSpanTable(labelParas, 'center')], section(), makeCtx());
      const tables = pages.map(retainedTableOn).filter((table) => table !== undefined);
      const restartPieces = tables.flatMap((table) => table.fragment.rows)
        .filter((row) => row.logicalRowIndex === 0);
      expect(restartPieces.map((row) => [
        row.fragmentIndex,
        row.cells[0]?.verticalMerge,
        row.cells[0]?.vAlign,
      ])).toEqual([
        [0, 'restart', 'center'],
        [1, 'restart', 'center'],
      ]);
      expect(restartPieces.flatMap((row) => sourceBlockIndices(row.cells[0])))
        .toEqual(Array.from({ length: labelParas }, (_v, index) => index));
      for (const table of tables) {
        expect(table.placed.heightPt).toBeLessThanOrEqual(100 + 1e-6);
      }
    }
  });

  it('derives page-local ownership when a semantic continue cell starts a fragment', () => {
    const label = mkCell([shortPara('L0')], true);
    const content = mkCell([shortPara('c0')], null);
    const tallContinue: DocTableRow = {
      cells: [
        mkCell([para({}) as CellElement], false),
        mkCell([para({ text: 'あ'.repeat(24), fontSize: 20 }) as CellElement], null),
      ],
      rowHeight: null, rowHeightRule: 'auto', isHeader: false,
    } as unknown as DocTableRow;
    const t: DocTable = {
      colWidths: [40, 120],
      rows: [
        { cells: [label, content], rowHeight: null, rowHeightRule: 'auto', isHeader: false } as unknown as DocTableRow,
        tallContinue,
      ],
      borders: emptyBorders(),
      cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
      jc: 'left', layout: 'fixed',
    } as unknown as DocTable;
    const pages = computePages(
      [para(), para(), para(), para(), { type: 'table', ...t } as BodyElement],
      section(),
      makeCtx(),
    );
    const first = retainedTableOn(pages[0]);
    const second = retainedTableOn(pages[1]);
    expect(first?.fragment.rows.map((row) => row.logicalRowIndex)).toEqual([0]);
    expect(first?.fragment.rows[0]?.cells[0]?.verticalMerge).toBe('restart');
    expect(second?.fragment.rows.map((row) => row.logicalRowIndex)).toEqual([1]);
    expect(second?.fragment.rows[0]?.cells[0]?.verticalMerge).toBe('continue');
    expect(second?.fragment.rows[0]?.cells[0]?.visualMergeOwnership).toBe('continuation');
    expect(second?.fragment.rows[0]?.cells[1]?.contentRanges).toEqual([
      { kind: 'whole', blockIndex: 0 },
    ]);
  });
});

describe('computePages — line-boundary splitting + widowControl (C1: §17.3.1.44)', () => {
  // contentW = 160, glyph advance = fontPx; at 20px → 8 chars/line. 48 chars → 6 lines.
  // content height = 100 → 5 lines (100px) fit per page.
  const sixLineText = 'あ'.repeat(48);

  it('avoids a widow: a single trailing line is not stranded on the next page (default widowControl on)', () => {
    const pages = computePages([para({ text: sixLineText })], section(), makeCtx());
    expect(pages.length).toBe(2);
    // Greedy fit is 5 lines on page 1; widowControl pulls one down so ≥2 carry over.
    expect(sliceOf(pages[0][0])).toEqual({ start: 0, end: 4 });
    expect(sliceOf(pages[1][0])).toEqual({ start: 4, end: 6 });
  });

  it('honors w:widowControl="off": the trailing single line is allowed (matches sample-9)', () => {
    const pages = computePages([para({ text: sixLineText, widowControl: false })], section(), makeCtx());
    expect(pages.length).toBe(2);
    expect(sliceOf(pages[0][0])).toEqual({ start: 0, end: 5 }); // greedy 5 lines
    expect(sliceOf(pages[1][0])).toEqual({ start: 5, end: 6 }); // lone widow line allowed
  });
});

describe('computePages — anchored wrap-shape float exclusion (B: §20.4.2.16)', () => {
  // A topAndBottom anchored text-box SHAPE must reserve a float band exactly
  // like an anchored image does, so body text in following paragraphs flows
  // BELOW it instead of overlapping. Geometry: page content height =
  // 140 - 20 - 20 = 100pt. The shape is anchored at the first paragraph's top
  // (≈ marginTop = 20pt) with height 50 ⇒ band y∈[20,70]. The following text
  // paragraph is 2 lines × 20pt = 40pt.
  //
  // Without the shape registering a float (the bug): the shape paragraph's mark
  // (20pt: y20→40) + the 2-line text (40pt: y40→80) all fit ⇒ 1 page.
  // With the float (the fix): the mark flows below the band (y70→90) and the
  // text is pushed under the band, so its 2 lines (≥ y90 → ≈130) overflow the
  // 100pt content area ⇒ 2 pages.
  it('pushes following text below a topAndBottom shape (shape registers a float)', () => {
    const body = [
      paraWith([shapeRun({ widthPt: 160, heightPt: 50, wrapMode: 'topAndBottom' })]),
      para({ text: 'あ'.repeat(16), fontSize: 20 }), // 160/20 = 8 chars/line → 2 lines
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(2);
    // The text paragraph (with its float-displaced first line) is pushed to the
    // second page; page 1 holds only the shape-anchoring paragraph.
    const onPage2 = pages[1].some(
      (el) => el.type === 'paragraph' &&
        (el as unknown as DocParagraph).runs.some((r) => r.type === 'text'),
    );
    expect(onPage2).toBe(true);
  });

  it('does NOT reserve a band for a wrapNone shape (no float, text stays on one page)', () => {
    // Same geometry but wrapMode:'none' ⇒ no exclusion rect; everything fits on
    // one page. Guards against over-registering floats for non-wrapping shapes.
    const body = [
      paraWith([shapeRun({ widthPt: 160, heightPt: 50, wrapMode: 'none' })]),
      para({ text: 'あ'.repeat(16), fontSize: 20 }),
    ];
    const pages = computePages(body, section(), makeCtx());
    expect(pages.length).toBe(1);
  });

  it('does not reserve flow height for a front non-callout wrapNone shape before a following inline image', () => {
    const imagePara = paraWith([inlineImageRun(80, 30)]);
    const body = [
      paraWith([shapeRun({ widthPt: 160, heightPt: 80, wrapMode: 'none' })]),
      imagePara,
    ];

    const pages = computePages(body, section(), makeCtx());
    expect(pages).toHaveLength(1);
    expect(pages[0].some((el) =>
      el.type === 'paragraph' &&
      (el as unknown as DocParagraph).runs.some((r) => r.type === 'image'),
    )).toBe(true);
  });

  it('does not let a wrapNone callout shape force a page break', () => {
    const imagePara = paraWith([inlineImageRun(80, 30)]);
    const callout = {
      ...shapeRun({ widthPt: 160, heightPt: 80, wrapMode: 'none' }),
      presetGeometry: 'accentBorderCallout2',
    } as DocRun;
    const body = [
      paraWith([callout]),
      imagePara,
    ];

    const pages = computePages(body, section(), makeCtx());
    expect(pages).toHaveLength(1);
  });

  it('does not move an anchor-only wrapNone callout to the next page to fit its box', () => {
    const calloutAnchor = paraWith([
      {
        ...shapeRun({ widthPt: 160, heightPt: 50, wrapMode: 'none', anchorYFromPara: true }),
        presetGeometry: 'accentBorderCallout2',
      } as DocRun,
    ]);
    const body = [
      ...Array.from({ length: 4 }, () => para()),
      calloutAnchor,
    ];

    const pages = computePages(body, section(), makeCtx());
    expect(pages).toHaveLength(1);
    const occurrence = paragraphOccurrenceWithRun(
      pages[0],
      (calloutAnchor as unknown as DocParagraph).runs[0]!,
    );
    expect(occurrence).toBeDefined();
    expect(occurrence).not.toBe(calloutAnchor);
    expect(occurrence ? bodyFragmentFor(occurrence)?.fragment.source.path[0] : undefined).toBe(4);
  });

  it('does not count a front wrapNone shape as flow height when grouping following inline images', () => {
    const smallImage = paraWith([inlineImageRun(80, 10)]);
    const photo = paraWith([inlineImageRun(80, 30)]);
    const body = [
      paraWith([shapeRun({ widthPt: 160, heightPt: 60, wrapMode: 'none' })]),
      smallImage,
      para(),
      photo,
    ];

    const pages = computePages(body, section(), makeCtx());
    expect(pages).toHaveLength(1);
    const smallImageOccurrence = paragraphOccurrenceWithRun(
      pages[0],
      (smallImage as unknown as DocParagraph).runs[0]!,
    );
    expect(smallImageOccurrence).toBeDefined();
    expect(smallImageOccurrence).not.toBe(smallImage);
    expect(smallImageOccurrence
      ? bodyFragmentFor(smallImageOccurrence)?.fragment.source.path[0]
      : undefined).toBe(1);
    expect(pages[0].some((el) =>
      el.type === 'paragraph' &&
      (el as unknown as DocParagraph).runs.some((r) => r.type === 'image'),
    )).toBe(true);
  });

  it('moves an inline-image paragraph to a fresh page when it no longer fits', () => {
    const photo = paraWith([inlineImageRun(80, 30)]);
    const body = [
      ...Array.from({ length: 4 }, () => para()),
      photo,
    ];

    const pages = computePages(body, section(), makeCtx());
    expect(pages).toHaveLength(2);
    expect(pages[0]).not.toContain(photo as PaginatedBodyElement);
    expect(pages[1].some((el) =>
      el.type === 'paragraph' &&
      (el as unknown as DocParagraph).runs.some((r) => r.type === 'image'),
    )).toBe(true);
  });
});

describe('computePages — paragraph anchor before explicit page break (§20.4.3.5 + §17.3.1.20)', () => {
  it('keeps a paragraph-anchored wrapNone shape on the pre-break page when a hard page break follows', () => {
    // Four empty paragraphs consume 80pt of the 100pt body. The following
    // wrapNone callout is 50pt tall from its paragraph top, so the generic
    // keep-on-page float rule would normally relocate it to a fresh page. But
    // the source immediately follows this anchor paragraph with a hard page
    // break: Word keeps the pre-break anchor on the pre-break page, then starts
    // the next paragraph after the authored break. sample-33 has this exact
    // shape after the "当日の板書" photo.
    const anchorPara = paraWith([
      shapeRun({ widthPt: 80, heightPt: 50, wrapMode: 'none', anchorYFromPara: true }),
    ]);
    const body = [
      ...Array.from({ length: 4 }, () => para()),
      anchorPara,
      pageBreak(),
      para({ text: 'after', fontSize: 20 }),
    ];

    const pages = computePages(body, section(), makeCtx());
    expect(pages).toHaveLength(2);
    expect(pages[0]).toContain(anchorPara as PaginatedBodyElement);
    expect(textOf(pages[1][0])).toBe('after');
  });

  it('does not push an anchor-only pre-break paragraph to a new page just for its empty mark', () => {
    const anchorPara = paraWith([
      shapeRun({ widthPt: 80, heightPt: 50, wrapMode: 'none', anchorYFromPara: true }),
    ]);
    const body = [
      ...Array.from({ length: 5 }, () => para()),
      anchorPara,
      pageBreak(),
      para({ text: 'after', fontSize: 20 }),
    ];

    const pages = computePages(body, section(), makeCtx());
    expect(pages).toHaveLength(2);
    expect(pages[0]).toContain(anchorPara as PaginatedBodyElement);
    expect(textOf(pages[1][0])).toBe('after');
  });

  it('moves a preceding image with its pre-break callout when the pair only fits fresh', () => {
    const imagePara = paraWith([inlineImageRun(80, 40)]);
    const anchorPara = paraWith([
      shapeRun({ widthPt: 80, heightPt: 50, wrapMode: 'none', anchorYFromPara: true }),
    ]);
    const body = [
      ...Array.from({ length: 3 }, () => para()),
      imagePara,
      anchorPara,
      pageBreak(),
      para({ text: 'after', fontSize: 20 }),
    ];

    const pages = computePages(body, section(), makeCtx());
    expect(pages).toHaveLength(3);
    const imageRun = (imagePara as unknown as DocParagraph).runs[0]!;
    expect(paragraphOccurrenceWithRun(pages[0], imageRun)).toBeUndefined();
    const imageOccurrence = paragraphOccurrenceWithRun(pages[1], imageRun);
    expect(imageOccurrence).toBeDefined();
    expect(imageOccurrence).not.toBe(imagePara);
    expect(imageOccurrence
      ? bodyFragmentFor(imageOccurrence)?.fragment.source.path[0]
      : undefined).toBe(3);
    expect(pages[1]).toContain(anchorPara as PaginatedBodyElement);
    expect(textOf(pages[2][0])).toBe('after');
  });
});

// ===== ECMA-376 §17.6.4 multi-column sections =====

describe('computeColumns — geometry (§17.6.4)', () => {
  it('count<=1 / no columns → one full-width column', () => {
    const cols = computeColumns(section());
    // contentW = 200 - 20 - 20 = 160; left = marginLeft = 20.
    expect(cols).toEqual([{ xPt: 20, wPt: 160 }]);
    // Explicit count:1 behaves the same.
    expect(computeColumns(section({ columns: { count: 1, spacePt: 36, equalWidth: true, sep: false, cols: [] } })))
      .toEqual([{ xPt: 20, wPt: 160 }]);
  });

  it('2 equal columns with a given space → correct x/w', () => {
    const cols = computeColumns(section({
      columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] },
    }));
    // colW = (160 - 1*20)/2 = 70. col0 at 20, col1 at 20 + 70 + 20 = 110.
    expect(cols).toEqual([
      { xPt: 20, wPt: 70 },
      { xPt: 110, wPt: 70 },
    ]);
  });

  it('3 equal columns tile the content band', () => {
    const cols = computeColumns(section({
      columns: { count: 3, spacePt: 10, equalWidth: true, sep: false, cols: [] },
    }));
    // colW = (160 - 2*10)/3 = 46.6667. Starts: 20, 76.6667, 133.3333.
    expect(cols).toHaveLength(3);
    expect(cols[0].xPt).toBeCloseTo(20, 6);
    expect(cols[0].wPt).toBeCloseTo(140 / 3, 6);
    expect(cols[1].xPt).toBeCloseTo(20 + 140 / 3 + 10, 6);
    expect(cols[2].xPt).toBeCloseTo(20 + 2 * (140 / 3 + 10), 6);
  });

  it('explicit <w:col> widths are used verbatim', () => {
    const cols = computeColumns(section({
      columns: {
        count: 2,
        spacePt: 0,
        equalWidth: false,
        sep: false,
        cols: [
          { widthPt: 100, spacePt: 12 },
          { widthPt: 48, spacePt: 0 },
        ],
      },
    }));
    // col0 at marginLeft=20, w=100; col1 at 20 + 100 + 12 = 132, w=48.
    expect(cols).toEqual([
      { xPt: 20, wPt: 100 },
      { xPt: 132, wPt: 48 },
    ]);
  });
});

describe('computePages — newspaper column flow (§17.6.4)', () => {
  const twoCol = (): SectionProps =>
    section({ columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] } });

  // Geometry: page content height = 140 - 40 = 100; each single-line para = 20px
  // ⇒ 5 lines fill a column. colW = (160-20)/2 = 70.

  it('overflowing paragraphs flow into column 1 on the SAME page (pages.length unchanged)', () => {
    // 7 single-line paragraphs: 5 fill column 0, 2 spill into column 1 — still 1 page.
    const body = Array.from({ length: 7 }, (_, i) => para({ text: `p${i}`, fontSize: 20 }));
    const pages = computePages(body, twoCol(), makeCtx());
    expect(pages.length).toBe(1);
    expect(pages[0]).toHaveLength(7);
    // First 5 in column 0, last 2 in column 1.
    expect(pages[0].slice(0, 5).map(colOf)).toEqual([0, 0, 0, 0, 0]);
    expect(pages[0].slice(5).map(colOf)).toEqual([1, 1]);
  });

  it('fills both columns before starting a new page', () => {
    // 11 single-line paras: col0=5, col1=5, then the 11th starts page 2 col0.
    const body = Array.from({ length: 11 }, (_, i) => para({ text: `p${i}`, fontSize: 20 }));
    const pages = computePages(body, twoCol(), makeCtx());
    expect(pages.length).toBe(2);
    expect(pages[0]).toHaveLength(10);
    expect(pages[0].slice(0, 5).every((el) => colOf(el) === 0)).toBe(true);
    expect(pages[0].slice(5).every((el) => colOf(el) === 1)).toBe(true);
    expect(pages[1]).toHaveLength(1);
    expect(colOf(pages[1][0])).toBe(0); // new page resets to column 0
  });

  it('measures at the column width: a paragraph wraps to MORE lines than at full width', () => {
    // 6 CJK chars. Full width (160 → 8/line) = 1 line. Column width (70 → 3/line) = 2 lines.
    const sixChars = 'あ'.repeat(6);
    // Single-column baseline: one line, easily 1 page.
    const single = computePages([para({ text: sixChars, fontSize: 20 })], section(), makeCtx());
    expect(single.length).toBe(1);
    expect(sliceOf(single[0][0])).toBeUndefined(); // not split — fits on one line

    // Two-column: 70/20 = 3 chars/line ⇒ the same paragraph needs 2 lines. Fill
    // column 0 with 4 single-line paras (4 lines), so only 1 line is left in
    // column 0; the 2-line paragraph cannot fit there and is pushed to column 1.
    const body = [
      ...Array.from({ length: 4 }, (_, i) => para({ text: `p${i}`, fontSize: 20 })),
      para({ text: sixChars, fontSize: 20 }),
    ];
    const pages = computePages(body, twoCol(), makeCtx());
    expect(pages.length).toBe(1);
    // The 2-line paragraph (proof it wrapped to >1 line at column width) moved to
    // column 1 because only one line of room remained in column 0.
    const wrapped = pages[0].find((el) => textOf(el) === sixChars);
    expect(wrapped).toBeDefined();
    expect(colOf(wrapped as PaginatedBodyElement)).toBe(1);
  });

  it('a long paragraph splits across columns of the same page', () => {
    // 18 CJK chars at column width (3/line) = 6 lines. Column holds 5 lines, so it
    // splits: 5 lines in column 0, 1 line in column 1 (widowControl off to keep the
    // greedy split deterministic) — both on page 1.
    const body = [para({ text: 'あ'.repeat(18), fontSize: 20, widowControl: false })];
    const pages = computePages(body, twoCol(), makeCtx());
    expect(pages.length).toBe(1);
    expect(pages[0]).toHaveLength(2);
    expect(sliceOf(pages[0][0])).toEqual({ start: 0, end: 5 });
    expect(colOf(pages[0][0])).toBe(0);
    expect(sliceOf(pages[0][1])).toEqual({ start: 5, end: 6 });
    expect(colOf(pages[0][1])).toBe(1);
  });
});

describe('computePages — explicit column break (§17.3.1.20)', () => {
  const twoCol = (): SectionProps =>
    section({ columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] } });

  it('a column break forces the next paragraph into column 1 (same page)', () => {
    const body = [
      para({ text: 'a', fontSize: 20 }),
      colBreak(),
      para({ text: 'b', fontSize: 20 }),
    ];
    const pages = computePages(body, twoCol(), makeCtx());
    expect(pages.length).toBe(1);
    // The break is consumed (not emitted as a body element); two paragraphs remain.
    const paras = pages[0].filter((el) => el.type === 'paragraph');
    expect(paras).toHaveLength(2);
    expect(colOf(paras[0])).toBe(0);
    expect(colOf(paras[1])).toBe(1);
  });

  it('a column break in the LAST column moves to the next page (column 0)', () => {
    const body = [
      para({ text: 'a', fontSize: 20 }),
      colBreak(), // → column 1
      para({ text: 'b', fontSize: 20 }),
      colBreak(), // last column → page 2, column 0
      para({ text: 'c', fontSize: 20 }),
    ];
    const pages = computePages(body, twoCol(), makeCtx());
    expect(pages.length).toBe(2);
    expect(colOf(pages[0].filter((e) => e.type === 'paragraph')[0])).toBe(0); // a
    expect(colOf(pages[0].filter((e) => e.type === 'paragraph')[1])).toBe(1); // b
    expect(textOf(pages[1][0])).toBe('c');
    expect(colOf(pages[1][0])).toBe(0); // c on page 2, column 0
  });

  it('ignores a trailing column break when no following content exists', () => {
    const body = [
      para({ text: 'a', fontSize: 20 }),
      colBreak(),
      para({ text: 'b', fontSize: 20 }),
      colBreak(),
    ];
    const pages = computePages(body, twoCol(), makeCtx());
    expect(pages.length).toBe(1);
    const paras = pages[0].filter((e) => e.type === 'paragraph');
    expect(paras.map(textOf)).toEqual(['a', 'b']);
    expect(colOf(paras[0])).toBe(0);
    expect(colOf(paras[1])).toBe(1);
  });
});

describe('computePages — single-column regression (§17.6.4 absent)', () => {
  it('a single-column section paginates identically (no colIndex > 0, full width)', () => {
    // Reuse the empty-paragraph relocation scenario: 7 empty paras, 5/page.
    const body = Array.from({ length: 7 }, () => para());
    const withCols = computePages(body, section(), makeCtx());
    expect(withCols.length).toBe(2);
    expect(withCols[0].length).toBe(5);
    expect(withCols[1].length).toBe(2);
    // Every placed element is column 0 (or untagged) — never a higher column.
    for (const page of withCols) {
      for (const el of page) expect(colOf(el) ?? 0).toBe(0);
    }
  });

  it('a wrapping paragraph splits the same way single-column as before', () => {
    // 48 CJK chars at full width (8/line) = 6 lines; 5 fit per page; widowControl
    // off ⇒ greedy 5 + 1. Identical to the pre-column behavior (guards no regression).
    const pages = computePages(
      [para({ text: 'あ'.repeat(48), fontSize: 20, widowControl: false })],
      section(),
      makeCtx(),
    );
    expect(pages.length).toBe(2);
    expect(sliceOf(pages[0][0])).toEqual({ start: 0, end: 5 });
    expect(sliceOf(pages[1][0])).toEqual({ start: 5, end: 6 });
    expect(colOf(pages[0][0]) ?? 0).toBe(0);
    expect(colOf(pages[1][0]) ?? 0).toBe(0);
  });
});

describe('computePages — per-section columns (§17.6.4, regression)', () => {
  // The sample-5 shape: section 1 is single-column (no <w:cols num>), only the
  // FINAL section is 2-column. Section 1's content MUST lay out in ONE full-width
  // column — not in the final section's 2-column grid. `section.columns` carries
  // the FINAL (body-level) section's columns; each mid-body section's columns ride
  // on its SectionBreak marker (here None ⇒ single column).
  const colGeomOf = (el: PaginatedBodyElement): { xPt: number; wPt: number }[] | undefined =>
    (el as { colGeom?: { xPt: number; wPt: number }[] }).colGeom;

  const finalTwoCol = (): SectionProps =>
    section({ columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] } });

  it("section 1 lays out in ONE full-width column even though the FINAL section is 2-col", () => {
    // 8 CJK chars: full width (160 → 8/line) = 1 line; the buggy 2-col path (70px →
    // 3/line) would wrap it to 3 lines. Section 1 = this para, ended by a nextPage
    // SectionBreak (cols None). Final section = one 2-col para.
    const body: BodyElement[] = [
      para({ text: 'あ'.repeat(8), fontSize: 20 }),
      sectionBreak('nextPage', null),
      para({ text: 'final', fontSize: 20 }),
    ];
    const pages = computePages(body, finalTwoCol(), makeCtx());
    // SectionBreak is consumed (not emitted); section 1 on page 1, final on page 2.
    expect(pages.length).toBe(2);
    const sec1Para = pages[0][0];
    // Full-width single column ⇒ NOT split (still one line) and colGeom length 1.
    expect(sliceOf(sec1Para)).toBeUndefined();
    expect(colOf(sec1Para) ?? 0).toBe(0);
    expect(colGeomOf(sec1Para)).toEqual([{ xPt: 20, wPt: 160 }]);
    // The FINAL section's paragraph gets the 2-column geometry.
    const finalPara = pages[1][0];
    expect(colGeomOf(finalPara)).toEqual([
      { xPt: 20, wPt: 70 },
      { xPt: 110, wPt: 70 },
    ]);
  });

  it('section 1 fills the full column height (5 lines/page) before paginating, not the half-width 2-col height', () => {
    // 5 single-line paragraphs fill a full-width column exactly (content height 100
    // / 20px = 5). All land on page 1, column 0, full-width geometry. Under the bug
    // they would have been squeezed into 2-col width and the flow would differ.
    const body: BodyElement[] = [
      ...Array.from({ length: 5 }, (_, i) => para({ text: `s${i}`, fontSize: 20 })),
      sectionBreak('nextPage', null),
      para({ text: 'final', fontSize: 20 }),
    ];
    const pages = computePages(body, finalTwoCol(), makeCtx());
    expect(pages[0]).toHaveLength(5);
    for (const el of pages[0]) {
      expect(colOf(el) ?? 0).toBe(0);
      expect(colGeomOf(el)).toEqual([{ xPt: 20, wPt: 160 }]);
    }
  });

  it('a continuous section break keeps both single-column sections on the SAME page, stacked', () => {
    // ECMA-376 §17.6.22: a section's `<w:type>` describes how THAT section begins
    // relative to the previous one, and lives on the sectPr that ENDS the section
    // (the NEXT marker in document order). So the A→B break is governed by section
    // B's type — the kind of the marker that ENDS B (idx 3) — and the B→final break
    // by the final (body) section's start type. Here B is continuous (stacks below
    // A on page 1) and the final 2-col section is nextPage (default) → page 2. Both
    // A and B are single-column full width and must NOT reset to the page top.
    const body: BodyElement[] = [
      para({ text: 'A', fontSize: 20 }),
      sectionBreak('continuous', null), // ends section A (A's own start type)
      para({ text: 'B', fontSize: 20 }),
      sectionBreak('continuous', null), // ends section B ⇒ B begins continuously (A→B same page)
      para({ text: 'final', fontSize: 20 }),
    ];
    const pages = computePages(body, finalTwoCol(), makeCtx());
    expect(pages.length).toBe(2);
    // A and B share page 1 (B continuous), final on page 2 (final section nextPage).
    expect(pages[0].map(textOf)).toEqual(['A', 'B']);
    expect(pages[1].map(textOf)).toEqual(['final']);
    for (const el of pages[0]) {
      expect(colGeomOf(el)).toEqual([{ xPt: 20, wPt: 160 }]);
    }
  });

  it('the break INTO a section is governed by that section start type, not the prior marker (§17.6.22)', () => {
    // Regression for the off-by-one: a title section whose own sectPr is nextPage
    // (the spec default) followed by a continuous body section must NOT page-break —
    // the boundary is governed by the FOLLOWING (body) section's start type. Mirrors
    // sample-13, where Word keeps the 2-col body on the title page.
    const body: BodyElement[] = [
      para({ text: 'title', fontSize: 20 }),
      sectionBreak('nextPage', null), // ends the title section (its own start type)
      para({ text: 'body', fontSize: 20 }),
    ];
    const sec = { ...finalTwoCol(), columns: null, sectionStart: 'continuous' };
    const pages = computePages(body, sec, makeCtx());
    expect(pages.length).toBe(1);
    expect(pages[0].map(textOf)).toEqual(['title', 'body']);
  });

  it('single-section 2-col docs are UNCHANGED (no SectionBreak markers ⇒ whole body uses section.columns)', () => {
    // Guard: with NO section breaks, columns = section.columns for the whole body,
    // identical to the pre-fix global-columns behavior (sample-10's case).
    const body = Array.from({ length: 7 }, (_, i) => para({ text: `p${i}`, fontSize: 20 }));
    const pages = computePages(body, finalTwoCol(), makeCtx());
    // 5 fill col0, 2 spill into col1 — one page (the existing newspaper-flow test).
    expect(pages.length).toBe(1);
    expect(pages[0].slice(0, 5).map(colOf)).toEqual([0, 0, 0, 0, 0]);
    expect(pages[0].slice(5).map(colOf)).toEqual([1, 1]);
    // Every element carries the same 2-col geometry (the body-level section's).
    for (const el of pages[0]) {
      expect(colGeomOf(el)).toEqual([
        { xPt: 20, wPt: 70 },
        { xPt: 110, wPt: 70 },
      ]);
    }
  });
});

describe('computePages — newspaper column balancing (§17.6.4, non-final continuous sections)', () => {
  const twoColSpec = { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] };

  it('balances a SHORT non-final 2-col section across both columns (not greedy col0)', () => {
    // section(): content height 100, content width 160, 2-col ⇒ col width 70.
    // 6 single-line paras × 20px = 120px total. Balanced = 60px/col = 3 paras each.
    // Greedy (the bug) would put 5 in col0 (fills to 100px) and 1 in col1.
    // The 2-col section ENDS with a continuous break ⇒ non-final ⇒ balanced.
    const body: BodyElement[] = [
      ...Array.from({ length: 6 }, (_, i) => para({ text: `p${i}`, fontSize: 20 })),
      sectionBreak('continuous', twoColSpec),
      para({ text: 'after', fontSize: 20 }),
    ];
    const pages = computePages(body, section(), makeCtx()); // final section = 1-col
    const colByText: Record<string, number | undefined> = {};
    for (const e of pages[0]) {
      const t = textOf(e);
      if (t.startsWith('p')) colByText[t] = colOf(e);
    }
    // Balanced: p0–p2 in col0, p3–p5 in col1.
    expect([colByText.p0, colByText.p1, colByText.p2]).toEqual([0, 0, 0]);
    expect([colByText.p3, colByText.p4, colByText.p5]).toEqual([1, 1, 1]);
  });

  it('splits a LONG paragraph across balanced columns at the LINE level (not packed whole into col0)', () => {
    // section(): content height 100, 2-col ⇒ col width 70, ~3 CJK chars/line at
    // 20px, 20px/line. A 24-char paragraph = 8 lines = 160px. The 2-col section
    // ends with a continuous break ⇒ balanced ⇒ balanceColH = 160/2 = 80px = 4
    // lines. The paragraph must SPLIT — ~4 lines in col0, the rest in col1 — not
    // pack all 8 lines into col0 (the sample-12 p.2 bug: a long first paragraph
    // left column 0 full and column 1 nearly empty).
    const body: BodyElement[] = [
      para({ text: 'あ'.repeat(24), fontSize: 20 }),
      sectionBreak('continuous', twoColSpec),
      para({ text: 'x', fontSize: 20 }),
    ];
    const pages = computePages(body, section(), makeCtx()); // final section = 1-col
    const slices = pages[0].filter((e) => e.type === 'paragraph' && sliceOf(e));
    const col0 = slices.filter((e) => (colOf(e) ?? 0) === 0);
    const col1 = slices.filter((e) => colOf(e) === 1);
    // One slice per column — the paragraph spans the balance boundary.
    expect(col0).toHaveLength(1);
    expect(col1).toHaveLength(1);
    const s0 = sliceOf(col0[0]) as { start: number; end: number };
    const s1 = sliceOf(col1[0]) as { start: number; end: number };
    expect(s0.start).toBe(0);
    expect(s0.end).toBe(4); // balance target = 4 lines
    expect(s1).toEqual({ start: 4, end: 8 }); // remainder in col1
  });

  it('moves a keepLines paragraph WHOLE to balance (does not split it across columns)', () => {
    // A keepLines paragraph (§17.3.1.14) cannot be split, so balancing falls back
    // to the whole-paragraph move. Short para (1 line) fills col0; the 3-line
    // keepLines paragraph exceeds the balance target (4 lines total ⇒ 2-line
    // target) and moves WHOLE into col1 — intact, no lineSlice.
    const body: BodyElement[] = [
      para({ text: 'x', fontSize: 20 }),
      para({ text: 'あ'.repeat(9), fontSize: 20, keepLines: true }), // 9 / 3 = 3 lines
      sectionBreak('continuous', twoColSpec),
      para({ text: 'after', fontSize: 20 }),
    ];
    const pages = computePages(body, section(), makeCtx());
    const keep = pages[0].find((e) => textOf(e).startsWith('あ'));
    expect(keep).toBeDefined();
    expect(colOf(keep as PaginatedBodyElement) ?? 0).toBe(1); // moved to col1
    expect(sliceOf(keep as PaginatedBodyElement)).toBeUndefined(); // whole, not split
  });

  it('does NOT balance the FINAL (body) section — it fills col0 greedily', () => {
    // Same 6 paras, but now the 2-col geometry is the body-level (final) section
    // with NO terminating break. Word leaves the final section greedy (the user's
    // observation: the last page packs the left column). Greedy: col0 holds 5
    // (100px), col1 holds 1.
    const body: BodyElement[] = Array.from({ length: 6 }, (_, i) => para({ text: `p${i}`, fontSize: 20 }));
    const pages = computePages(body, section({ columns: twoColSpec }), makeCtx());
    const col0 = pages[0].filter((e) => colOf(e) === 0).map(textOf);
    expect(col0).toEqual(['p0', 'p1', 'p2', 'p3', 'p4']); // greedy fill, NOT balanced
  });
});

describe('computePages — keepNext at a balanced column boundary (§17.3.1.15)', () => {
  const twoColSpec = { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] };

  // section(): content height 100, content width 160, 2-col ⇒ col width 70. Every
  // single-line para is 20px tall. 4 single-line paras ⇒ section total 80px ⇒
  // balanced target balanceColH = 80/2 = 40px = 2 lines/column. So col0 fits p0
  // (0..20) then a candidate at y=20 whose own line ends exactly AT the 40px
  // target — it fits alone, but adding its keepNext successor (another 20px) would
  // push the pair past the target.

  it('moves a keepNext paragraph WHOLE to the next column so it stays with its successor', () => {
    // p1 has keepNext: at y=20 it fits the balance target alone (20+20=40) but
    // p1+p2 (40+20=60) exceeds it. Pre-fix the balance break ignored keepNext, so
    // p1 stayed in col0 (0..40) and p2 spilled into col1 — orphaning the heading.
    const body: BodyElement[] = [
      para({ text: 'p0', fontSize: 20 }),
      para({ text: 'p1', fontSize: 20, keepNext: true }),
      para({ text: 'p2', fontSize: 20 }),
      para({ text: 'p3', fontSize: 20 }),
      sectionBreak('continuous', twoColSpec),
      para({ text: 'after', fontSize: 20 }),
    ];
    const pages = computePages(body, section(), makeCtx());
    const colByText: Record<string, number | undefined> = {};
    for (const e of pages[0]) {
      const t = textOf(e);
      if (t.startsWith('p')) colByText[t] = colOf(e) ?? 0;
    }
    // keepNext p1 and its successor p2 land in the SAME column (col1), not split.
    expect(colByText.p1).toBe(1);
    expect(colByText.p2).toBe(1);
    expect(colByText.p1).toBe(colByText.p2);
  });

  it('moves a keepNext paragraph WHOLE when its successor is a TABLE (§17.11 keep-with-next block)', () => {
    // Same geometry, but the block kept-with p1 is a 1-row 20px table. Its height
    // folds into the section total (80px ⇒ target 40px) and into needNext, so p1
    // must move to col1 rather than orphan above the table.
    const body: BodyElement[] = [
      para({ text: 'p0', fontSize: 20 }),
      para({ text: 'p1', fontSize: 20, keepNext: true }),
      fixedTable([20]),
      para({ text: 'p3', fontSize: 20 }),
      sectionBreak('continuous', twoColSpec),
      para({ text: 'after', fontSize: 20 }),
    ];
    const pages = computePages(body, section(), makeCtx());
    const p1 = pages[0].find((e) => textOf(e) === 'p1');
    const tbl = pages[0].find((e) => e.type === 'table');
    expect(p1).toBeDefined();
    expect(tbl).toBeDefined();
    expect(colOf(p1 as PaginatedBodyElement) ?? 0).toBe(1);
    // The table follows p1 in the SAME column.
    expect(colOf(tbl as PaginatedBodyElement) ?? 0).toBe(1);
  });

  it('does NOT move a paragraph WITHOUT keepNext (unchanged greedy-to-target fill)', () => {
    // Identical layout but p1 has no keepNext: it fills col0 up to the target
    // (0..40 with p0) and p2 spills into col1 as before — the fix must not perturb
    // ordinary balancing.
    const body: BodyElement[] = [
      para({ text: 'p0', fontSize: 20 }),
      para({ text: 'p1', fontSize: 20 }),
      para({ text: 'p2', fontSize: 20 }),
      para({ text: 'p3', fontSize: 20 }),
      sectionBreak('continuous', twoColSpec),
      para({ text: 'after', fontSize: 20 }),
    ];
    const pages = computePages(body, section(), makeCtx());
    const colByText: Record<string, number | undefined> = {};
    for (const e of pages[0]) {
      const t = textOf(e);
      if (t.startsWith('p')) colByText[t] = colOf(e) ?? 0;
    }
    expect(colByText.p0).toBe(0);
    expect(colByText.p1).toBe(0); // stays in col0 (no keepNext)
    expect(colByText.p2).toBe(1);
  });

  it('does NOT move a keepNext paragraph when the pair is taller than a balanced column (no infinite send)', () => {
    // p1 keepNext (1 line) is followed by a 3-line successor. section total: p0(1)
    // + p1(1) + p2(3) + p3(1) = 6 lines = 120px ⇒ target 60px = 3 lines. p1 at
    // y=20: p1 alone (40) ≤ target, but p1+p2 = 20+20+60 = 100px > 60px target —
    // the pair can never fit one balanced column, so moving p1 forward would loop.
    // p1 must stay put (its successor breaks normally), exactly as pre-fix.
    const body: BodyElement[] = [
      para({ text: 'p0', fontSize: 20 }),
      para({ text: 'p1', fontSize: 20, keepNext: true }),
      para({ text: 'あ'.repeat(9), fontSize: 20 }), // 9/3 = 3 lines = 60px
      para({ text: 'p3', fontSize: 20 }),
      sectionBreak('continuous', twoColSpec),
      para({ text: 'after', fontSize: 20 }),
    ];
    const pages = computePages(body, section(), makeCtx());
    const p1 = pages[0].find((e) => textOf(e) === 'p1');
    expect(p1).toBeDefined();
    expect(colOf(p1 as PaginatedBodyElement) ?? 0).toBe(0); // stays in col0
  });
});

describe('computePages — vanished (hidden) empty-paragraph mark collapse (§17.3.2.41 / §17.3.1.29)', () => {
  // ECMA-376 §17.3.1.29: a paragraph's mark has run properties. §17.3.2.41
  // `w:vanish` hides a run in the normal/print view (hidden-text off, the view a
  // Word PDF export renders). An INKLESS paragraph whose MARK is vanished is
  // therefore not displayed at all: it must contribute no mark line and no
  // paragraph spacing — the mark analogue of the parser stripping hidden runs.
  // sample-28 (issue #868): a run of empty vanished ListParagraphs otherwise
  // reserved ~156px and forced one extra page (24 vs Word's 23).
  //
  // Content height = pageHeight 140 − margins 40 = 100pt; fontSize 20 → each
  // single line / empty mark = 20pt (stub font box 0.8/0.2 em, no grid/spacing),
  // so exactly five single-line paragraphs fill one page.
  const visible = () => [
    para({ text: 'a', fontSize: 20 }),
    para({ text: 'b', fontSize: 20 }),
    para({ text: 'c', fontSize: 20 }),
    para({ text: 'd', fontSize: 20 }),
    para({ text: 'e', fontSize: 20 }),
  ];

  it('collapses inkless vanished-mark paragraphs to zero height (identical to their absence)', () => {
    const reference = computePages(
      [...visible().slice(0, 4), visible()[4]], // 5 visible paragraphs, no empties
      section(),
      makeCtx(),
    );
    const withVanish = computePages(
      [
        ...visible().slice(0, 4),
        para({ markVanish: true, spaceAfter: 8, fontSize: 20 }),
        para({ markVanish: true, spaceAfter: 8, fontSize: 20 }),
        para({ markVanish: true, spaceAfter: 8, fontSize: 20 }),
        visible()[4],
      ],
      section(),
      makeCtx(),
    );
    // The three vanished empties consume nothing: the fifth visible paragraph
    // lands on the same (single) page it would occupy without them.
    expect(reference.length).toBe(1);
    expect(withVanish.length).toBe(reference.length);
    // The fifth visible paragraph ('e') is on the first (only) page.
    expect(withVanish[0].some((el) => textOf(el) === 'e')).toBe(true);
  });

  it('control: ORDINARY empty paragraphs (mark not vanished) still reserve height and add a page', () => {
    const withOrdinaryEmpties = computePages(
      [
        ...visible().slice(0, 4),
        para({ spaceAfter: 8, fontSize: 20 }),
        para({ spaceAfter: 8, fontSize: 20 }),
        para({ spaceAfter: 8, fontSize: 20 }),
        visible()[4],
      ],
      section(),
      makeCtx(),
    );
    // Not a tautology: three non-vanished empties push the fifth visible past the
    // first page.
    expect(withOrdinaryEmpties.length).toBeGreaterThan(1);
  });
});
