import { beforeAll, describe, expect, it } from 'vitest';
import { layoutDocument } from '../document-layout.js';
import { createLayoutServices } from '../layout-runtime.js';
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  SectionProps,
} from '../types.js';
import { stableFingerprint } from './fingerprint.js';
import { createLayoutServicesRuntimeView } from './runtime-state.js';
import type { TextLayoutService, TextShapeRequest } from './text.js';
import type {
  FlowBlockPlacement,
  LayoutServices,
  ParagraphLayout,
  TableCellLayout,
  TableLayout,
  TableLayoutInput,
  TableRowLayout,
} from './types.js';
import { measureTableCellBlockFlowHeightPt } from './table.js';

function measuringContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText(text: string) {
      const fontSize = Number.parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...text].length * fontSize * 0.5,
        fontBoundingBoxAscent: fontSize * 0.8,
        fontBoundingBoxDescent: fontSize * 0.2,
        actualBoundingBoxAscent: fontSize * 0.8,
        actualBoundingBoxDescent: fontSize * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {}, createLinearGradient() {
      return { addColorStop() {} };
    },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000000', strokeStyle: '#000000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  } as unknown as CanvasRenderingContext2D;
}

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    getContext() { return measuringContext(); }
  };
});

function emptyBorders() {
  return { top: null, right: null, bottom: null, left: null, insideH: null, insideV: null };
}

function paragraph(text: string): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null,
      fontFamily: 'Times New Roman', fontFamilyEastAsia: '', isLink: false,
      background: null, vertAlign: null, hyperlink: null,
    } as DocParagraph['runs'][number]],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function cell(content: CellElement[], overrides: Partial<DocTableCell> = {}): DocTableCell {
  return {
    content, colSpan: 1, vMerge: null, borders: emptyBorders(), background: null,
    vAlign: 'top', widthPt: null, ...overrides,
  } as unknown as DocTableCell;
}

function textCell(text: string): DocTableCell {
  return cell([{ type: 'paragraph', ...paragraph(text) } as CellElement]);
}

function row(cells: DocTableCell[]): DocTableRow {
  return { cells, rowHeight: null, rowHeightRule: 'auto', isHeader: false } as DocTableRow;
}

function table(rows: DocTableRow[], columnWidthsPt: number[]): DocTable {
  return {
    type: 'table', rows, colWidths: columnWidthsPt, borders: emptyBorders(),
    cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
    jc: 'left', layout: 'fixed',
  } as unknown as DocTable;
}

function document(body: BodyElement[]): DocxDocumentModel {
  const section = {
    pageWidth: 220, pageHeight: 300,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage', columns: null,
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
    footnotes: [],
  } as unknown as DocxDocumentModel;
}

function countingServices(model: DocxDocumentModel): {
  services: LayoutServices;
  finalAcquisitions: ReadonlyMap<string, number>;
} {
  const base = createLayoutServices(model, { measureContext: measuringContext() });
  const counts = new Map<string, number>();
  const text: TextLayoutService = Object.freeze({
    fingerprint: base.text.fingerprint,
    localMetrics: base.text.localMetrics,
    resolve: (request: Parameters<TextLayoutService['resolve']>[0]) => base.text.resolve(request),
    shape(request: TextShapeRequest) {
      if (request.clusterGeometry === true) {
        counts.set(request.text, (counts.get(request.text) ?? 0) + 1);
      }
      return base.text.shape(request);
    },
  });
  return { services: createLayoutServicesRuntimeView(base, { text }), finalAcquisitions: counts };
}

function retainedParagraph(
  id: string,
  path: readonly number[],
  advancePt: number,
): ParagraphLayout {
  return {
    kind: 'paragraph', id,
    source: { story: 'body', storyInstance: 'body', path: [...path] },
    flowDomainId: 'body', ordinaryFlow: true,
    flowBounds: { xPt: 0, yPt: 0, widthPt: 20, heightPt: advancePt },
    inkBounds: { xPt: 0, yPt: 0, widthPt: 10, heightPt: advancePt },
    advancePt, spacing: { beforePt: 0, afterPt: 0 }, contextualSpacing: false,
    lines: [], borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
  };
}

type BorderInput = Readonly<{
  widthPt: number;
  color: string;
  authoredStyle: string;
}>;

type EdgeInputs = Readonly<{
  top: BorderInput | null;
  right: BorderInput | null;
  bottom: BorderInput | null;
  left: BorderInput | null;
  insideH: BorderInput | null;
  insideV: BorderInput | null;
}>;

const noBorderInputs: EdgeInputs = Object.freeze({
  top: null, right: null, bottom: null, left: null, insideH: null, insideV: null,
});

function placement(): FlowBlockPlacement {
  return {
    container: {
      id: 'body', kind: 'body', bounds: { xPt: 10, yPt: 20, widthPt: 100, heightPt: 200 },
    },
    cursor: { xPt: 10, yPt: 20 },
    availableBounds: { xPt: 10, yPt: 20, widthPt: 100, heightPt: 200 },
  };
}

function tableInput(
  rows: readonly unknown[],
  columnWidthsPt: readonly number[] = [40, 60],
  borders: EdgeInputs = noBorderInputs,
): TableLayoutInput {
  return {
    kind: 'table', id: 'table-0',
    source: { story: 'body', storyInstance: 'body', path: [0] },
    flowDomainId: 'body', ordinaryFlow: true,
    alignment: 'left', indentPt: 0, bidiVisual: false,
    columnWidthsPt, borders,
    // Source rows follow each other unless a test says otherwise.
    rows: rows.map((row, index) => ({ logicalRowIndex: index, ...(row as object) })),
  } as unknown as TableLayoutInput;
}

function cellInput(
  id: string,
  columnStart: number,
  blocks: readonly ParagraphLayout[],
  overrides: Readonly<Record<string, unknown>> = {},
) {
  return {
    id,
    source: { story: 'body', storyInstance: 'body', path: [0, Number(id.at(-1) ?? 0)] },
    columnStart, columnSpan: 1, verticalMerge: 'none',
    margins: { topPt: 0, rightPt: 0, bottomPt: 0, leftPt: 0 },
    vAlign: 'top', borders: noBorderInputs,
    blocks: blocks.map((layout) => ({ layout })),
    ...overrides,
  };
}

describe('retained table layout', () => {
  it('retains complete compound outer-frame membership for measurement-free paint', async () => {
    const compound = {
      widthPt: 1.5,
      color: '#000000',
      authoredStyle: 'thinThickSmallGap',
    };
    const rows = [{
      id: 'row-0',
      source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      heightPt: 20,
      heightRule: 'exact',
      cells: [cellInput('cell-0', 0, [])],
    }];
    const model = document([]);
    const services = createLayoutServices(model, { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const result = layoutTable(tableInput(rows, [100], {
      ...noBorderInputs,
      top: compound,
      right: compound,
      bottom: compound,
      left: compound,
    }), placement(), services);
    const layout = result.layout as TableLayout;

    // The frame joins the rules' centre lines: the top rule lies inside the
    // table below its edge, the bottom rule inside above the extended end.
    expect(layout.compoundBorderFrames).toEqual([{
      bounds: { xPt: 10, yPt: 20.75, widthPt: 100, heightPt: 20 },
      border: {
        authoredStyle: 'thinThickSmallGap',
        color: '#000000',
        widthPt: 1.5,
        style: 'compound',
      },
      segmentIndexes: [0, 1, 2, 3],
    }]);
  });

  it('includes each collapsed horizontal rule once, in the row below it, and the bottom rule last', async () => {
    // WORD_COLLAPSED_BORDER_ROW_TRACK_FOOTPRINT (Windows Word #1904 controls).
    const single = { widthPt: 0.5, color: '#000000', authoredStyle: 'single' };
    const rows = Array.from({ length: 4 }, (_, rowIndex) => ({
      id: `row-${rowIndex}`,
      source: { story: 'body', storyInstance: 'body', path: [0, rowIndex] },
      heightPt: 20.4,
      heightRule: 'atLeast',
      cells: [cellInput(`cell-${rowIndex}`, 0, [], {
        borders: { ...noBorderInputs, top: single, bottom: single },
      })],
    }));
    const model = document([]);
    const services = createLayoutServices(model, { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const result = layoutTable(tableInput(rows, [100]), placement(), services);
    const layout = result.layout as TableLayout;

    expect(layout.rows.map((rowLayout) => rowLayout.advancePt)).toEqual([
      20.9, 20.9, 20.9, 21.4,
    ]);
    expect(layout.advancePt).toBeCloseTo(84.1, 8);
  });

  it('keeps an exact row, which holds its top rule, and adds only the table bottom rule', async () => {
    const single = { widthPt: 0.5, color: '#000000', authoredStyle: 'single' };
    const input = tableInput([{
      id: 'row-0',
      source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      heightPt: 20.4,
      heightRule: 'exact',
      cells: [cellInput('cell-0', 0, [], {
        borders: { ...noBorderInputs, top: single, bottom: single },
      })],
    }], [100]);
    const model = document([]);
    const services = createLayoutServices(model, { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const result = layoutTable(input, placement(), services);

    expect((result.layout as TableLayout).rows[0]?.advancePt).toBeCloseTo(20.9, 8);
  });

  it('meets each vertical rule with the adjacent top and bottom rules', async () => {
    // Nonuniform horizontal rules: each vertical edge starts and ends at the
    // centre line of the rule in its own column, not the table-wide widest.
    const thin = { widthPt: 0.5, color: '#000000', authoredStyle: 'single' };
    const thick = { widthPt: 6, color: '#000000', authoredStyle: 'single' };
    const input = tableInput([{
      id: 'row-0',
      source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      heightPt: 20,
      heightRule: 'exact',
      cells: [
        cellInput('cell-0', 0, [], { borders: { ...noBorderInputs, top: thin, bottom: thin, left: thin } }),
        cellInput('cell-1', 1, [], { borders: { ...noBorderInputs, top: thick, bottom: thick, right: thick } }),
      ],
    }], [40, 60]);
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');
    const { layout } = layoutTable(input, placement(), services);
    const edge = (name: 'left' | 'right') => layout.borders.find((border) => border.edge === name)!;
    const topPt = layout.flowBounds.yPt;
    const bottomPt = topPt + layout.advancePt;
    expect([edge('left').from.yPt, edge('left').to.yPt]).toEqual([topPt + 0.25, bottomPt - 0.25]);
    expect([edge('right').from.yPt, edge('right').to.yPt]).toEqual([topPt + 3, bottomPt - 3]);
  });

  it('insets a prefix candidate below a merge by the merge terminal rule', async () => {
    // A vertical merge ending above the candidate ends in its 6pt bottom rule
    // (on its last cell); the 1pt insideH is only the fallback. The prefix
    // must resolve that boundary as the whole table does.
    const owner = { widthPt: 6, color: '#000000', authoredStyle: 'single' };
    const inside = { widthPt: 1, color: '#000000', authoredStyle: 'single' };
    const rowAt = (index: number, cell: Record<string, unknown>) => ({
      id: `row-${index}`,
      source: { story: 'body', storyInstance: 'body', path: [0, index] },
      logicalRowIndex: index,
      heightPt: 20,
      heightRule: 'atLeast',
      cells: [cell],
    });
    const rows = [
      rowAt(0, cellInput('cell-0', 0, [], { verticalMerge: 'restart' })),
      rowAt(1, cellInput('cell-1', 0, [], { verticalMerge: 'continue', borders: { ...noBorderInputs, bottom: owner } })),
      rowAt(2, cellInput('cell-2', 0, [])),
    ];
    const input = tableInput(rows, [100], { ...noBorderInputs, insideH: inside });
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable, tablePrefixTracks } = await import('./table.js');
    const full = layoutTable(input, placement(), services).layout as TableLayout;
    const prefix = tablePrefixTracks(input).row(input.rows.slice(0, 2), input.rows[2]!, placement())!;
    const inset = (rowLayout: TableLayout['rows'][number]) =>
      rowLayout.cells[0]!.contentBounds.yPt - rowLayout.cells[0]!.flowBounds.yPt;
    expect(inset(full.rows[2]!)).toBeCloseTo(6, 9);
    expect(inset(prefix)).toBeCloseTo(6, 9);
  });

  it('acquires each ordinary and nested cell paragraph final geometry once', () => {
    const nested = table([row([textCell('nested-only')])], [80]);
    const outer = table([row([
      textCell('outer-only'),
      cell([{ type: 'table', ...nested } as CellElement]),
    ])], [90, 90]);
    const model = document([outer as BodyElement]);
    const { services, finalAcquisitions } = countingServices(model);

    const result = layoutDocument(model, services);

    expect(result.pages).toHaveLength(1);
    expect({
      ordinary: finalAcquisitions.get('outer-only'),
      nested: finalAcquisitions.get('nested-only'),
    }).toEqual({ ordinary: 1, nested: 1 });
  });

  it('feeds private row height, spacing, and cell margins into retained geometry', () => {
    const sourceCell = {
      ...textCell('private-format'),
      __tableCellLayout: {
        preferredWidth: null,
        margins: {
          top: { kind: 'dxa', value: '0' },
          right: { kind: 'dxa', value: '120' },
          bottom: { kind: 'dxa', value: '40' },
          left: { kind: 'dxa', value: '80' },
          start: null, end: null,
        },
      },
    };
    const source = {
      ...table([{
        ...row([sourceCell as unknown as DocTableCell]),
        __tableRowLayout: {
          height: { value: '200', rule: 'exact', ruleAuthored: true },
          justification: null,
          beforeWidth: null, afterWidth: null,
          cellSpacing: null, exception: null,
        },
      } as unknown as DocTableRow], [40]),
      __tableLayout: {
        effectiveStyleId: null,
        grid: { authored: true, columns: [{ width: '800' }], requiredColumnCount: 1 },
        preferredWidth: null, layout: { kind: 'fixed' },
        cellSpacing: { kind: 'dxa', value: '40' }, cellMargins: null,
      },
    } as unknown as DocTable;
    const model = document([source as BodyElement]);
    const services = createLayoutServices(model, { measureContext: measuringContext() });

    const result = layoutDocument(model, services);
    const retained = result.pages.flatMap((page) => page.layers.body)
      .find((node): node is TableLayout => node.kind === 'table');

    expect(retained).toBeDefined();
    expect(retained?.rows[0]?.advancePt).toBe(12);
    expect(retained?.rows[0]?.cells[0]?.flowBounds).toMatchObject({ widthPt: 36, heightPt: 8 });
    expect(retained?.rows[0]?.cells[0]?.contentBounds).toMatchObject({ widthPt: 26, heightPt: 6 });
    expect(
      (retained?.rows[0]?.cells[0]?.contentBounds.xPt ?? 0) - (retained?.flowBounds.xPt ?? 0),
    ).toBe(6);
  });

  it('returns one immutable parser-independent retained tree', async () => {
    const paragraphLayout = retainedParagraph('p0', [0, 0, 0], 8);
    const input = tableInput([{
      id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      heightPt: null, heightRule: 'auto',
      cells: [cellInput('cell-0', 0, [paragraphLayout], {
        margins: { topPt: 1, rightPt: 2, bottomPt: 3, leftPt: 4 },
        background: { color: '#abcdef' },
      })],
    }], [40]);
    const model = document([]);
    const services = createLayoutServices(model, { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const result = layoutTable(input, placement(), services);
    const layout = result.layout as TableLayout;
    const before = stableFingerprint('table-layout', layout);

    expect(layout.source).toEqual({ story: 'body', storyInstance: 'body', path: [0] });
    expect(layout.source).not.toBe(input.source);
    expect(Object.isFrozen(layout)).toBe(true);
    expect(Object.isFrozen(layout.columnWidthsPt)).toBe(true);
    expect(Object.isFrozen(layout.rows)).toBe(true);
    expect(Object.isFrozen(layout.rows[0])).toBe(true);
    expect(Object.isFrozen(layout.rows[0]?.cells)).toBe(true);
    expect(Object.isFrozen(layout.rows[0]?.cells[0])).toBe(true);
    expect(Object.isFrozen(layout.rows[0]?.cells[0]?.blocks)).toBe(true);
    expect(Object.isFrozen(layout.borders)).toBe(true);
    expect(() => structuredClone(layout)).not.toThrow();
    expect(Object.getPrototypeOf(structuredClone(layout))).toBe(Object.prototype);

    const mutable = input as unknown as {
      columnWidthsPt: number[];
      rows: Array<{ cells: Array<{ margins: { topPt: number } }> }>;
      source: { path: number[] };
    };
    mutable.columnWidthsPt[0] = 999;
    mutable.rows[0]!.cells[0]!.margins.topPt = 999;
    mutable.source.path[0] = 999;
    expect(stableFingerprint('table-layout', layout)).toBe(before);
  });

  it('derives row and cell geometry from retained children, margins, floors, and vMerge', async () => {
    const merged = retainedParagraph('merged', [0, 0, 0], 25);
    const topSide = retainedParagraph('top-side', [0, 0, 1], 8);
    const bottomSide = retainedParagraph('bottom-side', [0, 1, 1], 6);
    const input = tableInput([
      {
        id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
        heightPt: null, heightRule: 'auto',
        cells: [
          cellInput('cell-0', 0, [merged], {
            verticalMerge: 'restart',
            margins: { topPt: 2, rightPt: 0, bottomPt: 3, leftPt: 0 },
          }),
          cellInput('cell-1', 1, [topSide], {
            margins: { topPt: 1, rightPt: 5, bottomPt: 1, leftPt: 4 },
          }),
        ],
      },
      {
        id: 'row-1', source: { story: 'body', storyInstance: 'body', path: [0, 1] },
        heightPt: 12, heightRule: 'atLeast',
        cells: [
          cellInput('cell-2', 0, [], { verticalMerge: 'continue' }),
          cellInput('cell-3', 1, [bottomSide], {
            margins: { topPt: 1, rightPt: 5, bottomPt: 2, leftPt: 4 },
          }),
        ],
      },
    ]);
    const model = document([]);
    const services = createLayoutServices(model, { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const result = layoutTable(input, placement(), services);
    const layout = result.layout as TableLayout & {
      rows: readonly TableRowLayout[];
    };
    const firstRow = layout.rows[0] as TableRowLayout;
    const secondRow = layout.rows[1] as TableRowLayout;
    const topSideCell = firstRow.cells[1] as TableCellLayout;
    const continuation = secondRow.cells[0] as TableCellLayout;

    // The ordinary top-side cell establishes the auto row at 1 + 8 + 1 = 10pt.
    // The second row starts at its 12pt floor. The 30pt merged content exceeds
    // the 22pt span, so the deficit is assigned to the span's last row.
    expect(firstRow.flowBounds).toMatchObject({ xPt: 10, yPt: 20, widthPt: 100, heightPt: 10 });
    expect(secondRow.flowBounds).toMatchObject({ xPt: 10, yPt: 30, widthPt: 100, heightPt: 20 });
    expect(layout.advancePt).toBe(30);
    expect(result.nextCursor).toEqual({ xPt: 10, yPt: 50 });

    expect(topSideCell.contentBounds).toEqual({
      xPt: 54, yPt: 21, widthPt: 51, heightPt: 8,
    });
    expect(topSideCell.blocks).toEqual([
      expect.objectContaining({ layout: topSide, offsetPt: 1, advancePt: 8 }),
    ]);
    expect(continuation.verticalMerge).toBe('continue');
    expect(continuation.blocks).toEqual([]);
  });

  it('adds collapsed boundary footprints to non-exact row-track allocation', async () => {
    const single = { widthPt: 1, color: '#000000', authoredStyle: 'single' };
    const input = tableInput([
      {
        id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
        heightPt: 8, heightRule: 'atLeast',
        cells: [cellInput('cell-0', 0, [])],
      },
      {
        id: 'row-1', source: { story: 'body', storyInstance: 'body', path: [0, 1] },
        heightPt: 8, heightRule: 'atLeast',
        cells: [cellInput('cell-1', 0, [])],
      },
      {
        id: 'row-2', source: { story: 'body', storyInstance: 'body', path: [0, 2] },
        heightPt: 8, heightRule: 'exact',
        cells: [cellInput('cell-2', 0, [])],
      },
    ], [40], {
      ...noBorderInputs,
      top: single,
      insideH: single,
      bottom: single,
    });
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);

    // atLeast rows add their top rule; the exact row already holds its own
    // and the table's bottom rule extends it.
    expect(layout.rows.map((row) => row.advancePt)).toEqual([9, 9, 9]);
    expect(layout.advancePt).toBe(27);
  });

  it('keeps explicit auto content-sized and adds Word bottom padding to exact heights', async () => {
    const content = retainedParagraph('content', [0, 0, 0], 8);
    const input = tableInput([
      {
        id: 'row-auto', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
        heightPt: 40, heightRule: 'auto',
        cells: [cellInput('cell-0', 0, [content])],
      },
      {
        id: 'row-exact', source: { story: 'body', storyInstance: 'body', path: [0, 1] },
        heightPt: 10, heightRule: 'exact',
        cells: [cellInput('cell-1', 0, [retainedParagraph('overflow', [0, 1, 0], 30)], {
          margins: { topPt: 1, rightPt: 0, bottomPt: 3, leftPt: 0 },
        })],
      },
    ], [40]);
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);

    // ECMA-376 §17.4.80: explicit auto ignores @val and follows content.
    expect(layout.rows[0]?.advancePt).toBe(8);
    // [MS-OI29500] 2.1.180(d): Word adds the largest bottom cell padding to an
    // exact row, while excess content remains clipped rather than growing it.
    expect(layout.rows[1]?.advancePt).toBe(13);
    expect(layout.rows[1]?.contentHeightPt).toBe(34);
    expect(layout.rows[1]?.cells[0]?.clipBounds).toEqual({
      xPt: 10,
      yPt: layout.rows[1]?.cells[0]?.flowBounds.yPt,
      widthPt: 100,
      heightPt: layout.rows[1]?.cells[0]?.flowBounds.heightPt,
    });
  });

  it.each(['center', 'bottom'] as const)(
    'top-aligns overflowing exact-row content instead of moving its first line above the %s cell',
    async (vAlign) => {
      const overflow = retainedParagraph('overflow', [0, 0, 0], 30);
      const input = tableInput([{
        id: 'row-exact', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
        heightPt: 10, heightRule: 'exact',
        cells: [cellInput('cell-0', 0, [overflow], { vAlign })],
      }], [40]);
      const services = createLayoutServices(document([]), { measureContext: measuringContext() });
      const { layoutTable } = await import('./table.js');

      const { layout } = layoutTable(input, placement(), services);
      const cell = layout.rows[0]!.cells[0]!;

      // Word keeps the beginning of an over-height story at the top inset and
      // lets the exact-row clip discard its tail. A negative center/bottom
      // offset would instead discard the first line, reversing source order.
      expect(cell.blocks[0]?.offsetPt).toBe(0);
      expect(cell.clipBounds).toEqual({
        xPt: 10, yPt: 20, widthPt: 100, heightPt: 10,
      });
    },
  );

  it('satisfies a merged interval only through the latest growable row', async () => {
    const input = tableInput([
      {
        id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
        heightPt: 10, heightRule: 'auto',
        cells: [cellInput('cell-0', 0, [retainedParagraph('merged', [0, 0, 0], 30)], {
          verticalMerge: 'restart',
        })],
      },
      {
        id: 'row-1', source: { story: 'body', storyInstance: 'body', path: [0, 1] },
        heightPt: 8, heightRule: 'exact',
        cells: [cellInput('cell-1', 0, [], { verticalMerge: 'continue' })],
      },
      {
        id: 'row-2', source: { story: 'body', storyInstance: 'body', path: [0, 2] },
        heightPt: 6, heightRule: 'atLeast',
        cells: [cellInput('cell-2', 0, [], { verticalMerge: 'continue' })],
      },
    ], [40]);
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);

    // The explicit auto value is ignored, exact stays fixed, and the terminal
    // growable track receives the interval deficit: 0 + 8 + 22 = 30.
    expect(layout.rows.map((rowLayout) => rowLayout.advancePt)).toEqual([0, 8, 22]);
  });

  it('clips an overflowing merged owner when every row in its span is exact', async () => {
    const input = tableInput([
      {
        id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
        heightPt: 8, heightRule: 'exact',
        cells: [cellInput('cell-0', 0, [retainedParagraph('merged', [0, 0, 0], 30)], {
          verticalMerge: 'restart',
        })],
      },
      {
        id: 'row-1', source: { story: 'body', storyInstance: 'body', path: [0, 1] },
        heightPt: 7, heightRule: 'exact',
        cells: [cellInput('cell-1', 0, [], { verticalMerge: 'continue' })],
      },
    ], [40]);
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);
    const owner = layout.rows[0]!.cells[0]!;

    expect(layout.rows.map((rowLayout) => rowLayout.advancePt)).toEqual([8, 7]);
    expect(owner.flowBounds.heightPt).toBe(15);
    expect(owner.clipBounds).toEqual({ xPt: 10, yPt: 20, widthPt: 100, heightPt: 15 });
  });

  it('propagates translated child ink through cell, row, and table bounds', async () => {
    const child = {
      ...retainedParagraph('overhang', [0, 0, 0], 8),
      inkBounds: { xPt: -6, yPt: -2, widthPt: 30, heightPt: 12 },
    };
    const input = tableInput([{
      id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      heightPt: null, heightRule: 'auto',
      cells: [cellInput('cell-0', 0, [child])],
    }], [40]);
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);

    expect(layout.rows[0]!.cells[0]!.inkBounds)
      .toEqual({ xPt: 4, yPt: 18, widthPt: 46, heightPt: 12 });
    expect(layout.rows[0]!.inkBounds)
      .toEqual({ xPt: 4, yPt: 18, widthPt: 46, heightPt: 12 });
    expect(layout.inkBounds)
      .toEqual({ xPt: 4, yPt: 18, widthPt: 46, heightPt: 12 });
  });

  it('uses only the retained layoutInCell extent to grow an automatic row', async () => {
    const child = {
      ...retainedParagraph('cell-anchor', [0, 0, 0], 10),
      inkBounds: { xPt: -20, yPt: -20, widthPt: 200, heightPt: 200 },
      cellContainmentBounds: { xPt: 0, yPt: 0, widthPt: 30, heightPt: 80 },
    } as ParagraphLayout;
    const input = tableInput([{
      id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      heightPt: null, heightRule: 'auto',
      cells: [cellInput('cell-0', 0, [child])],
    }], [40]);
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);

    expect(measureTableCellBlockFlowHeightPt([{ layout: child, sourceBlockIndex: 0 }]))
      .toBe(80);
    expect(layout.rows[0]!.advancePt).toBe(80);
    expect(layout.rows[0]!.cells[0]!.flowBounds.heightPt).toBe(80);
  });

  it('includes retained layoutInCell geometry in bottom cell alignment', async () => {
    const child = {
      ...retainedParagraph('bottom-cell-anchor', [0, 0, 0], 10),
      cellContainmentBounds: { xPt: 0, yPt: 0, widthPt: 30, heightPt: 80 },
    } as ParagraphLayout;
    const input = tableInput([{
      id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      heightPt: null, heightRule: 'auto',
      cells: [cellInput('cell-0', 0, [child], { vAlign: 'bottom' })],
    }], [40]);
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);
    const cell = layout.rows[0]!.cells[0]!;
    const block = cell.blocks[0]!;
    const placedContainmentBottomPt = cell.flowBounds.yPt
      + block.offsetPt
      + child.cellContainmentBounds!.yPt
      + child.cellContainmentBounds!.heightPt
      - child.flowBounds.yPt;

    expect(placedContainmentBottomPt).toBeLessThanOrEqual(
      cell.flowBounds.yPt + cell.flowBounds.heightPt,
    );
  });

  it('grows and offsets a top-aligned cell for negative layoutInCell geometry', async () => {
    const child = {
      ...retainedParagraph('negative-top-cell-anchor', [0, 0, 0], 10),
      cellContainmentBounds: { xPt: 0, yPt: -20, widthPt: 30, heightPt: 60 },
    } as ParagraphLayout;
    const input = tableInput([{
      id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      heightPt: null, heightRule: 'auto',
      cells: [cellInput('cell-0', 0, [child], { vAlign: 'top' })],
    }], [40]);
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);
    const cell = layout.rows[0]!.cells[0]!;
    const block = cell.blocks[0]!;
    const placedContainmentTopPt = cell.flowBounds.yPt
      + block.offsetPt
      + child.cellContainmentBounds!.yPt
      - child.flowBounds.yPt;
    const placedContainmentBottomPt = placedContainmentTopPt
      + child.cellContainmentBounds!.heightPt;

    expect(cell.flowBounds.heightPt).toBe(60);
    expect(placedContainmentTopPt).toBeGreaterThanOrEqual(cell.flowBounds.yPt);
    expect(placedContainmentBottomPt).toBeLessThanOrEqual(
      cell.flowBounds.yPt + cell.flowBounds.heightPt,
    );
  });

  it('clips retained layoutInCell overflow instead of growing an exact row', async () => {
    const child = {
      ...retainedParagraph('exact-cell-anchor', [0, 0, 0], 10),
      cellContainmentBounds: { xPt: 0, yPt: 0, widthPt: 30, heightPt: 80 },
    } as ParagraphLayout;
    const input = tableInput([{
      id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      heightPt: 20, heightRule: 'exact',
      cells: [cellInput('cell-0', 0, [child])],
    }], [40]);
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);
    const cell = layout.rows[0]!.cells[0]!;

    expect(layout.rows[0]!.advancePt).toBe(20);
    expect(cell.flowBounds.heightPt).toBe(20);
    expect(cell.clipBounds).toEqual({ xPt: 10, yPt: 20, widthPt: 100, heightPt: 20 });
  });

  it('retains the opposing shared segment when the other candidate is nil', async () => {
    const nil = { widthPt: 1, color: '#000000', authoredStyle: 'nil' };
    const single = { widthPt: 1, color: '#000000', authoredStyle: 'single' };
    const input = tableInput([{
      id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      heightPt: 8, heightRule: 'exact', cellSpacingPt: 0,
      cells: [
        cellInput('cell-0', 0, [], { borders: { ...noBorderInputs, right: nil } }),
        cellInput('cell-1', 1, [], { borders: { ...noBorderInputs, left: single } }),
      ],
    }]);
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);

    expect(layout.borders.filter((border) => border.edge === 'between')).toEqual([
      expect.objectContaining({
        color: '#000000',
        widthPt: 1,
      }),
    ]);
  });

  it('treats a cell-level inside border as a cell conflict candidate', async () => {
    const strongInside = { widthPt: 2, color: '#111111', authoredStyle: 'thick' };
    const weakDirect = { widthPt: 0.125, color: '#222222', authoredStyle: 'single' };
    const input = tableInput([{
      id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      heightPt: 8, heightRule: 'exact',
      cells: [
        cellInput('cell-0', 0, [], {
          borders: { ...noBorderInputs, insideV: strongInside },
        }),
        cellInput('cell-1', 1, [], {
          borders: { ...noBorderInputs, left: weakDirect },
        }),
      ],
    }]);
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);

    expect(layout.borders.find((border) => border.edge === 'between'))
      .toMatchObject({ widthPt: 2, color: '#111111' });
  });

  it('keeps cell spacing inside the table width and between every cell border box', async () => {
    const input = tableInput([
      {
        id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
        heightPt: null, heightRule: 'auto', cellSpacingPt: 4,
        cells: [
          cellInput('cell-0', 0, [retainedParagraph('a', [0, 0, 0], 8)]),
          cellInput('cell-1', 1, [retainedParagraph('b', [0, 0, 1], 8)]),
        ],
      },
      {
        id: 'row-1', source: { story: 'body', storyInstance: 'body', path: [0, 1] },
        heightPt: null, heightRule: 'auto', cellSpacingPt: 4,
        cells: [
          cellInput('cell-2', 0, [retainedParagraph('c', [0, 1, 0], 8)]),
          cellInput('cell-3', 1, [retainedParagraph('d', [0, 1, 1], 8)]),
        ],
      },
    ]);
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);

    expect(layout.flowBounds).toEqual({ xPt: 10, yPt: 20, widthPt: 100, heightPt: 28 });
    expect(layout.rows.map((item) => item.advancePt)).toEqual([14, 14]);
    expect(layout.rows[0]?.cells.map((item) => item.flowBounds)).toEqual([
      { xPt: 14, yPt: 24, widthPt: 34, heightPt: 8 },
      { xPt: 52, yPt: 24, widthPt: 54, heightPt: 8 },
    ]);
    expect(layout.rows[1]?.cells.map((item) => item.flowBounds)).toEqual([
      { xPt: 14, yPt: 36, widthPt: 34, heightPt: 8 },
      { xPt: 52, yPt: 36, widthPt: 54, heightPt: 8 },
    ]);
  });

  it('retains both opposing cell borders when non-zero spacing disables conflict resolution', async () => {
    const red = { widthPt: 1, color: '#ff0000', authoredStyle: 'single' };
    const blue = { widthPt: 1, color: '#0000ff', authoredStyle: 'single' };
    const input = tableInput([{
      id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      heightPt: 8, heightRule: 'exact', cellSpacingPt: 4,
      cells: [
        cellInput('cell-0', 0, [], { borders: { ...noBorderInputs, right: red } }),
        cellInput('cell-1', 1, [], { borders: { ...noBorderInputs, left: blue } }),
      ],
    }]);
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);

    expect(layout.borders.map((border) => ({
      x: border.from.xPt,
      color: border.color,
    }))).toEqual(expect.arrayContaining([
      { x: 48, color: '#ff0000' },
      { x: 52, color: '#0000ff' },
    ]));
  });

  it('applies the Word insideV conflict deviation when cells are spaced', async () => {
    const tableInside = { widthPt: 4, color: '#ff0000', authoredStyle: 'thick' };
    const conditionalInside = { widthPt: 1, color: '#0000ff', authoredStyle: 'single' };
    const input = tableInput([{
      id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      heightPt: 8, heightRule: 'exact', cellSpacingPt: 4,
      cells: [
        cellInput('cell-0', 0, [], {
          borders: { ...noBorderInputs, insideV: conditionalInside },
        }),
        cellInput('cell-1', 1, [], {
          borders: { ...noBorderInputs, insideV: conditionalInside },
        }),
      ],
    }], [40, 60], { ...noBorderInputs, insideV: tableInside });
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);

    // [MS-OI29500] 2.1.138 keeps conditional insideV subject to the
    // tcBorders/tblBorders conflict even with spacing. A cell candidate wins
    // over the table candidate, so the shared-grid red rule is suppressed and
    // the two inset blue cell edges remain.
    expect(layout.borders.map((border) => ({
      x: border.from.xPt,
      color: border.color,
    }))).toEqual(expect.arrayContaining([
      { x: 48, color: '#0000ff' },
      { x: 52, color: '#0000ff' },
    ]));
    expect(layout.borders.some((border) => border.color === '#ff0000')).toBe(false);
  });

  it('applies the Word insideH conflict deviation when rows are spaced', async () => {
    const tableInside = { widthPt: 4, color: '#ff0000', authoredStyle: 'thick' };
    const conditionalInside = { widthPt: 1, color: '#0000ff', authoredStyle: 'single' };
    const rows = [0, 1].map((rowIndex) => ({
      id: `row-${rowIndex}`,
      source: { story: 'body' as const, storyInstance: 'body', path: [0, rowIndex] },
      heightPt: 12, heightRule: 'exact' as const, cellSpacingPt: 4,
      cells: [cellInput(`cell-${rowIndex}`, 0, [], {
        borders: { ...noBorderInputs, insideH: conditionalInside },
      })],
    }));
    const input = tableInput(rows, [40], { ...noBorderInputs, insideH: tableInside });
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);
    const blueRules = layout.borders.filter((border) => border.color === '#0000ff');

    expect(blueRules).toHaveLength(2);
    expect(new Set(blueRules.map((border) => border.from.yPt)).size).toBe(2);
    expect(layout.borders.some((border) => border.color === '#ff0000')).toBe(false);
  });

  it('does not draw a spaced table insideV border through a spanning cell', async () => {
    const tableInside = { widthPt: 1, color: '#ff0000', authoredStyle: 'single' };
    const input = tableInput([{
      id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      heightPt: 12, heightRule: 'exact', cellSpacingPt: 4,
      cells: [cellInput('cell-0', 0, [], { columnSpan: 2 })],
    }], [40, 60], { ...noBorderInputs, insideV: tableInside });
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);

    expect(layout.borders.filter((border) => border.color === '#ff0000')).toEqual([]);
  });

  it('keeps spaced table-property exception borders on the shared grid', async () => {
    const tableInside = { widthPt: 1, color: '#ff0000', authoredStyle: 'single' };
    const exceptionInside = { widthPt: 2, color: '#0000ff', authoredStyle: 'thick' };
    const input = tableInput([{
      id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      heightPt: 12, heightRule: 'exact', cellSpacingPt: 4,
      exceptionBorders: { ...noBorderInputs, insideV: exceptionInside },
      cells: [cellInput('cell-0', 0, []), cellInput('cell-1', 1, [])],
    }], [40, 60], { ...noBorderInputs, insideV: tableInside });
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);

    expect(layout.borders.filter((border) => (
      border.color === '#0000ff' || border.color === '#ff0000'
    ))).toEqual([
      expect.objectContaining({
        color: '#0000ff', from: { xPt: 50, yPt: 20 }, to: { xPt: 50, yPt: 32 },
      }),
    ]);
  });

  it('keeps spaced horizontal table-property exception borders on the shared grid', async () => {
    const tableInside = { widthPt: 1, color: '#ff0000', authoredStyle: 'single' };
    const exceptionInside = { widthPt: 2, color: '#0000ff', authoredStyle: 'thick' };
    const rows = [0, 1].map((rowIndex) => ({
      id: `row-${rowIndex}`,
      source: { story: 'body' as const, storyInstance: 'body', path: [0, rowIndex] },
      heightPt: 12, heightRule: 'exact' as const, cellSpacingPt: 4,
      exceptionBorders: { ...noBorderInputs, insideH: exceptionInside },
      cells: [cellInput(`cell-${rowIndex}`, 0, [])],
    }));
    const input = tableInput(rows, [100], { ...noBorderInputs, insideH: tableInside });
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);

    expect(layout.borders.filter((border) => (
      border.color === '#0000ff' || border.color === '#ff0000'
    ))).toEqual([
      expect.objectContaining({
        color: '#0000ff', from: { xPt: 10, yPt: 32 }, to: { xPt: 110, yPt: 32 },
      }),
    ]);
  });

  it('does not draw a spaced table insideH border through a vertical merge', async () => {
    const tableInside = { widthPt: 1, color: '#ff0000', authoredStyle: 'single' };
    const input = tableInput([
      {
        id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
        heightPt: 12, heightRule: 'exact', cellSpacingPt: 4,
        cells: [
          cellInput('cell-0', 0, [], { verticalMerge: 'restart' }),
          cellInput('cell-1', 1, []),
        ],
      },
      {
        id: 'row-1', source: { story: 'body', storyInstance: 'body', path: [0, 1] },
        heightPt: 12, heightRule: 'exact', cellSpacingPt: 4,
        cells: [
          cellInput('cell-2', 0, [], { verticalMerge: 'continue' }),
          cellInput('cell-3', 1, []),
        ],
      },
    ], [40, 60], { ...noBorderInputs, insideH: tableInside });
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);
    const inside = layout.borders.filter((border) => border.color === '#ff0000');

    expect(inside).toEqual([
      expect.objectContaining({ from: { xPt: 50, yPt: 32 }, to: { xPt: 110, yPt: 32 } }),
    ]);
  });

  it('retains both row borders when cell spacing begins at their shared boundary', async () => {
    const red = { widthPt: 1, color: '#ff0000', authoredStyle: 'single' };
    const blue = { widthPt: 1, color: '#0000ff', authoredStyle: 'single' };
    const input = tableInput([
      {
        id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
        heightPt: 8, heightRule: 'exact', cellSpacingPt: 0,
        cells: [cellInput('cell-0', 0, [], {
          borders: { ...noBorderInputs, bottom: red },
        })],
      },
      {
        id: 'row-1', source: { story: 'body', storyInstance: 'body', path: [0, 1] },
        heightPt: 8, heightRule: 'exact', cellSpacingPt: 4,
        cells: [cellInput('cell-1', 0, [], {
          borders: { ...noBorderInputs, top: blue },
        })],
      },
    ], [40]);
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);

    expect(layout.borders.map((border) => border.color)).toEqual(expect.arrayContaining([
      '#ff0000',
      '#0000ff',
    ]));
  });

  it('cascades a row table-property exception between cell and table borders', async () => {
    const weakTable = { widthPt: 1, color: '#aaaaaa', authoredStyle: 'single' };
    const rowException = { widthPt: 2, color: '#222222', authoredStyle: 'thick' };
    const cellOverride = { widthPt: 0.5, color: '#ff0000', authoredStyle: 'single' };
    const input = tableInput([{
      id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      heightPt: 8, heightRule: 'exact', cellSpacingPt: 0,
      exceptionBorders: { ...noBorderInputs, top: rowException },
      cells: [
        cellInput('cell-0', 0, []),
        cellInput('cell-1', 1, [], { borders: { ...noBorderInputs, top: cellOverride } }),
      ],
    }], [40, 60], { ...noBorderInputs, top: weakTable });
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);
    const top = layout.borders.filter((border) => border.edge === 'top');

    expect(top).toEqual([
      expect.objectContaining({ color: '#222222', widthPt: 2 }),
      expect.objectContaining({ color: '#ff0000', widthPt: 0.5 }),
    ]);
  });

  it('treats a Word cell border none as absent before row/table fallback', async () => {
    const tableTop = { widthPt: 1, color: '#aaaaaa', authoredStyle: 'single' };
    const exceptionTop = { widthPt: 2, color: '#222222', authoredStyle: 'thick' };
    const none = { widthPt: 0, color: '#000000', authoredStyle: 'none' };
    const input = tableInput([{
      id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      heightPt: 8, heightRule: 'exact', cellSpacingPt: 0,
      exceptionBorders: { ...noBorderInputs, top: exceptionTop },
      cells: [cellInput('cell-0', 0, [], {
        borders: { ...noBorderInputs, top: none },
      })],
    }], [40], { ...noBorderInputs, top: tableTop });
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);

    expect(layout.borders.filter((border) => border.edge === 'top')).toEqual([
      expect.objectContaining({ color: '#222222', widthPt: 2 }),
    ]);
  });

  it('places each row by its effective table-row alignment', async () => {
    const content = retainedParagraph('row', [0, 0, 0], 8);
    const input = {
      ...tableInput([
        {
          id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
          heightPt: null, heightRule: 'auto', cellSpacingPt: 0, exceptionBorders: null,
          alignment: 'center', indentPt: 0,
          cells: [cellInput('cell-0', 0, [content])],
        },
        {
          id: 'row-1', source: { story: 'body', storyInstance: 'body', path: [0, 1] },
          heightPt: null, heightRule: 'auto', cellSpacingPt: 0, exceptionBorders: null,
          alignment: 'left', indentPt: 0,
          cells: [cellInput('cell-1', 0, [content])],
        },
      ], [40]),
      alignment: 'center' as const,
    };
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);

    expect(layout.rows.map((item) => item.flowBounds.xPt)).toEqual([40, 10]);
    expect(layout.rows.map((item) => item.cells[0]?.flowBounds.xPt)).toEqual([40, 10]);
    expect(layout.flowBounds).toEqual({ xPt: 10, yPt: 20, widthPt: 70, heightPt: 16 });
  });

  it('conflicts borders only where differently aligned row edges physically overlap', async () => {
    const red = { widthPt: 1, color: '#ff0000', authoredStyle: 'single' };
    const blue = { widthPt: 1, color: '#0000ff', authoredStyle: 'single' };
    const input = tableInput([
      {
        id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
        heightPt: 8, heightRule: 'exact', cellSpacingPt: 0, exceptionBorders: null,
        alignment: 'center', indentPt: 0,
        cells: [cellInput('cell-0', 0, [], {
          borders: { ...noBorderInputs, bottom: red },
        })],
      },
      {
        id: 'row-1', source: { story: 'body', storyInstance: 'body', path: [0, 1] },
        heightPt: 8, heightRule: 'exact', cellSpacingPt: 0, exceptionBorders: null,
        alignment: 'left', indentPt: 0,
        cells: [cellInput('cell-1', 0, [], {
          borders: { ...noBorderInputs, top: blue },
        })],
      },
    ], [40]);
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);

    expect(layout.borders.filter((border) => border.edge === 'between')).toEqual([
      expect.objectContaining({
        color: '#0000ff', from: { xPt: 10, yPt: 28.5 }, to: { xPt: 40, yPt: 28.5 },
      }),
      expect.objectContaining({
        color: '#ff0000', from: { xPt: 40, yPt: 28.5 }, to: { xPt: 80, yPt: 28.5 },
      }),
    ]);
  });

  it('resolves physical overlap across different logical columns after a row shift', async () => {
    const red = { widthPt: 1, color: '#ff0000', authoredStyle: 'single' };
    const blue = { widthPt: 1, color: '#0000ff', authoredStyle: 'single' };
    const input = tableInput([
      {
        id: 'row-0', source: { story: 'body', storyInstance: 'body', path: [0, 0] },
        heightPt: 8, heightRule: 'exact', cellSpacingPt: 0, exceptionBorders: null,
        alignment: 'center', indentPt: 0,
        cells: [
          cellInput('cell-0', 0, [], { borders: { ...noBorderInputs, bottom: red } }),
          cellInput('cell-1', 1, [], { borders: { ...noBorderInputs, bottom: red } }),
        ],
      },
      {
        id: 'row-1', source: { story: 'body', storyInstance: 'body', path: [0, 1] },
        heightPt: 8, heightRule: 'exact', cellSpacingPt: 0, exceptionBorders: null,
        alignment: 'left', indentPt: 0,
        cells: [
          cellInput('cell-2', 0, [], { borders: { ...noBorderInputs, top: blue } }),
          cellInput('cell-3', 1, [], { borders: { ...noBorderInputs, top: blue } }),
        ],
      },
    ], [20, 20]);
    const services = createLayoutServices(document([]), { measureContext: measuringContext() });
    const { layoutTable } = await import('./table.js');

    const { layout } = layoutTable(input, placement(), services);

    expect(layout.borders.filter((border) => border.edge === 'between')).toEqual([
      expect.objectContaining({
        color: '#0000ff', from: { xPt: 10, yPt: 28.5 }, to: { xPt: 40, yPt: 28.5 },
      }),
      expect.objectContaining({
        color: '#ff0000', from: { xPt: 40, yPt: 28.5 }, to: { xPt: 80, yPt: 28.5 },
      }),
    ]);
  });
});
