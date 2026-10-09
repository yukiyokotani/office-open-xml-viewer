import { beforeAll, describe, expect, it } from 'vitest';
import { layoutDocument } from '../document-layout.js';
import { tableFormatInput } from '../parser-model.js';
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
import type { TableFragmentLayout } from './table-pagination.js';

function makeStubCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const context = {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText(text: string) {
      const fontSize = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
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
  };
  (context as unknown as { canvas: unknown }).canvas = { width: 2000, height: 2000 };
  return context as unknown as CanvasRenderingContext2D;
}

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    getContext() { return makeStubCtx(); }
  };
});

function emptyBorders() {
  return { top: null, right: null, bottom: null, left: null, insideH: null, insideV: null };
}

function singleBorder(width: number) {
  return { style: 'single', width, color: '#000000' } as const;
}

function paragraph(text: string, fontSize = 10): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: text.length === 0 ? [] : [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize, color: null, fontFamily: 'Times New Roman',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null,
      hyperlink: null,
    } as DocParagraph['runs'][number]],
    defaultFontSize: fontSize, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function cell(content: CellElement[], overrides: Partial<DocTableCell> = {}): DocTableCell {
  return {
    content, colSpan: 1, vMerge: null, borders: emptyBorders(),
    background: null, vAlign: 'top', widthPt: null,
    ...overrides,
  } as unknown as DocTableCell;
}

function textCell(text: string): DocTableCell {
  return cell([{ type: 'paragraph', ...paragraph(text) } as unknown as CellElement]);
}

function row(cells: DocTableCell[], overrides: Partial<DocTableRow> = {}): DocTableRow {
  return {
    cells, rowHeight: null, rowHeightRule: 'auto', isHeader: false,
    ...overrides,
  } as unknown as DocTableRow;
}

function table(
  rows: DocTableRow[],
  columnWidthsPt: number[],
  overrides: Partial<DocTable> = {},
): DocTable {
  return {
    type: 'table', rows, colWidths: columnWidthsPt, borders: emptyBorders(),
    cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
    jc: 'left', layout: 'fixed',
    ...overrides,
  } as unknown as DocTable;
}

function documentModel(
  body: BodyElement[],
  pageHeight: number,
  pageWidth = 200,
): DocxDocumentModel {
  return {
    section: {
      pageWidth, pageHeight,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
      sectionStart: 'nextPage', columns: null,
    } as SectionProps,
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
    footnotes: [],
  } as unknown as DocxDocumentModel;
}

function retainedTopLevelTables(layout: ReturnType<typeof layoutDocument>): TableFragmentLayout[] {
  return layout.pages.flatMap((page) => page.layers.body.flatMap((node) => {
    return node.kind === 'table'
      ? [node as TableFragmentLayout]
      : [];
  }));
}

function expectDeeplyFrozen(value: unknown, seen = new WeakSet<object>()): void {
  if (value === null || typeof value !== 'object' || seen.has(value)) return;
  seen.add(value);
  expect(Object.isFrozen(value)).toBe(true);
  for (const child of Object.values(value)) expectDeeplyFrozen(child, seen);
}

describe('retained table pagination contracts', () => {
  it('places a cell-owned grid frame in the page margin band at the cell insertion cursor', () => {
    for (const childWidth of [160, 260]) {
      const nested = table([row([textCell('grid')], { rowHeight: 20, rowHeightRule: 'exact' })],
        [childWidth], { widthPt: childWidth });
      Object.assign(nested, { __tableLayout: {
        effectiveStyleId: null, ordinaryFlow: false,
        grid: { authored: true, columns: [{ width: String(childWidth * 20) }], requiredColumnCount: 1 },
        preferredWidth: { kind: 'dxa', value: String(childWidth * 20) },
        layout: { kind: 'fixed' }, cellSpacing: null,
        cellFrame: { dropCap: 'none', lines: 1, wrap: 'around', hAnchor: 'margin',
          vAnchor: 'text', hRule: 'auto', hSpace: 9, vSpace: 0, xAlign: 'center', y: 12 },
      } });
      const following = { ...paragraph(''), spaceBefore: 35 };
      const host = table([row([cell([nested as CellElement, following as CellElement])],
        { rowHeight: 70, rowHeightRule: 'exact' })], [100]);
      const [fragment] = retainedTopLevelTables(layoutDocument(documentModel([host as BodyElement], 120)));
      const placements = fragment?.resolvedFloatingTables ?? [];
      expect(placements).toHaveLength(1);
      const placed = placements[0]!;
      const acquiredWidth = placed.child.columnWidthsPt.reduce((a, b) => a + b, 0);
      expect(acquiredWidth).toBe(childWidth);
      expect(placed.xPt).toBeCloseTo(50, 10);
      expect(placed.exclusionBounds.widthPt).toBe(118);
      expect(placed.yPt).toBeCloseTo(22, 10);
      expect(fragment?.rows[0]?.cells[0]?.blocks.some((block) => block.layout.kind === 'table')).toBe(false);
      expect(placed.child.rows[0]?.cells[0]?.blocks[0]?.layout.kind).toBe('paragraph');
      expect(nested.tblpPr).toBeUndefined();
    }
  });
  it('retains nested page-split geometry as clone-safe immutable parser-independent data', () => {
    const nested = table(
      Array.from({ length: 4 }, (_unused, index) => row(
        [textCell(`nested ${index}`)],
        { rowHeight: 30, rowHeightRule: 'exact' },
      )),
      [80],
    );
    const sourceCell = cell([{ type: 'table', ...nested } as unknown as CellElement], {
      background: '112233',
    });
    const source = table([row([sourceCell])], [120]);
    const layout = layoutDocument(documentModel([source as unknown as BodyElement], 80));
    const nestedFragments = retainedTopLevelTables(layout).flatMap((fragment) =>
      fragment.rows.flatMap((fragmentRow) => fragmentRow.cells.flatMap((fragmentCell) =>
        fragmentCell.blocks.flatMap((block) => block.layout.kind === 'table'
          ? [block.layout as TableFragmentLayout]
          : []),
      )),
    );

    expect(nestedFragments.length).toBeGreaterThan(1);
    for (const fragment of nestedFragments) {
      expect(() => structuredClone(fragment)).not.toThrow();
      expectDeeplyFrozen(fragment);
    }
    const fingerprint = JSON.stringify(nestedFragments);

    source.jc = 'right';
    sourceCell.background = 'ffffff';
    nested.rows[0]!.cells[0]!.vAlign = 'bottom';
    nested.rows[1]!.rowHeight = 99;

    expect(JSON.stringify(nestedFragments)).toBe(fingerprint);
  });

  it('isolates placements and widths across independent pagination runs', () => {
    const source = table([row([textCell('independent')])], [200], { widthPct: 5000 });
    const body = [source as unknown as BodyElement];
    const wideModel = documentModel(body, 100, 200);
    const narrowModel = {
      ...wideModel,
      section: { ...wideModel.section, pageWidth: 150 },
    } as DocxDocumentModel;

    const wide = layoutDocument(wideModel).pages[0]?.layers.body.find((node) => node.kind === 'table');
    if (wide?.kind !== 'table') throw new Error('expected wide table placement');
    const wideFingerprint = JSON.stringify(wide);

    const narrow = layoutDocument(narrowModel).pages[0]?.layers.body.find((node) => node.kind === 'table');
    if (narrow?.kind !== 'table') throw new Error('expected narrow table placement');

    expect(wide.flowBounds.widthPt).toBeCloseTo(180, 6);
    expect(narrow.flowBounds.widthPt).toBeCloseTo(130, 6);
    expect(wide.flowBounds.xPt).toBeCloseTo(10, 6);
    expect(narrow.flowBounds.xPt).toBeCloseTo(10, 6);
    expect(wide).not.toBe(narrow);
    expect(JSON.stringify(wide)).toBe(wideFingerprint);
    expect(source).not.toHaveProperty('tableColWidthsPt');
    expect(source).not.toHaveProperty('tableRowHeightsPt');
  });

  it('assigns a reused table the source path of its current document occurrence', () => {
    const source = table([row([textCell('reused')])], [120]);
    const firstModel = documentModel([source as unknown as BodyElement], 100);
    const secondModel = documentModel([
      paragraph('before') as unknown as BodyElement,
      source as unknown as BodyElement,
    ], 100);

    const [first] = retainedTopLevelTables(layoutDocument(firstModel));
    const [second] = retainedTopLevelTables(layoutDocument(secondModel));

    expect(first?.source.path[0]).toBe(0);
    expect(second?.source.path[0]).toBe(1);
    expect(second?.rows[0]?.cells[0]?.blocks[0]?.layout.source.path[0]).toBe(1);
  });

  it('keeps repeated table occurrences distinct across unequal body columns', () => {
    const source = table(
      [row([textCell('repeated')], { rowHeight: 30, rowHeightRule: 'exact' })],
      [180],
      { widthPct: 5000 },
    );
    const model = documentModel([
      source as unknown as BodyElement,
      { type: 'columnBreak' } as BodyElement,
      source as unknown as BodyElement,
    ], 100, 220);
    model.section = {
      ...model.section,
      columns: {
        count: 2,
        spacePt: 20,
        equalWidth: false,
        sep: false,
        cols: [
          { widthPt: 120, spacePt: 20 },
          { widthPt: 60, spacePt: 0 },
        ],
      },
    } as SectionProps;

    const layout = layoutDocument(model);
    const placements = layout.pages.flatMap((page) => page.layers.body.flatMap((node) => (
      node.kind === 'table' ? [{ page, node }] : []
    )));

    expect(placements).toHaveLength(2);
    expect(placements.map(({ node }) => node.source.path[0])).toEqual([0, 2]);
    expect(placements.map(({ page, node }) => page.sectionRegions
      .flatMap((region) => region.flowDomainIds)
      .indexOf(node.flowDomainId))).toEqual([0, 1]);
    expect(placements.map(({ node }) => node.flowBounds.widthPt)).toEqual([120, 60]);
    expect(placements[0]?.node).not.toBe(placements[1]?.node);
    expect(placements[0]?.node.source.path).not.toEqual(placements[1]?.node.source.path);
  });

  it('emits distinct retained paragraph envelopes for repeated source occurrences', () => {
    const source = paragraph('repeated paragraph');
    const layout = layoutDocument(documentModel([
      source as unknown as BodyElement,
      source as unknown as BodyElement,
    ], 100));
    const emitted = layout.pages.flatMap((page) => page.layers.body).filter(
      (node) => node.kind === 'paragraph',
    );

    expect(emitted).toHaveLength(2);
    expect(emitted[0]).not.toBe(emitted[1]);
    expect(emitted.map((node) => node.source.path[0])).toEqual([0, 1]);
    expect(emitted[0]?.flowBounds.yPt).toBeLessThan(emitted[1]?.flowBounds.yPt ?? 0);
  });

  it('keeps an earlier retained paragraph envelope stable across sessions', () => {
    const source = paragraph('independent paragraph session');
    const wideModel = documentModel([source as unknown as BodyElement], 100, 200);
    const narrowModel = documentModel([source as unknown as BodyElement], 100, 150);
    const firstElement = layoutDocument(wideModel).pages[0]?.layers.body.find(
      (node) => node.kind === 'paragraph',
    );
    if (!firstElement) throw new Error('expected the first retained paragraph occurrence');
    const firstFingerprint = JSON.stringify(firstElement);

    const secondElement = layoutDocument(narrowModel).pages[0]?.layers.body.find(
      (node) => node.kind === 'paragraph',
    );
    if (!secondElement) throw new Error('expected the second retained paragraph occurrence');

    expect(firstElement).not.toBe(secondElement);
    expect(firstElement.flowBounds.widthPt).toBeCloseTo(180, 6);
    expect(secondElement.flowBounds.widthPt).toBeCloseTo(130, 6);
    expect(JSON.stringify(firstElement)).toBe(firstFingerprint);
  });

  it('resolves page-local outer border ink for every auto-height slice', () => {
    const outer = singleBorder(4);
    const inside = singleBorder(1);
    const source = table(
      Array.from({ length: 3 }, () => row(
        [cell([], { marginTop: 10, marginBottom: 10 })],
        { rowHeight: null, rowHeightRule: 'auto' },
      )),
      [120],
      {
        borders: {
          top: outer, bottom: outer, left: null, right: null,
          insideH: inside, insideV: null,
        },
      },
    );

    expect(tableFormatInput(source).rows.map((format) => format.height?.rule ?? 'auto'))
      .toEqual(['auto', 'auto', 'auto']);

    const fragments = retainedTopLevelTables(
      layoutDocument(documentModel([source as unknown as BodyElement], 50)),
    );

    expect(fragments).toHaveLength(3);
    for (const fragment of fragments) {
      expect(fragment.rows).toHaveLength(1);
      expect(fragment.rows[0]?.heightPt).toBe(24);
      expect(fragment.advancePt).toBe(24);
      expect(fragment.borders.filter((border) => border.edge === 'top')).toEqual([
        expect.objectContaining({ widthPt: 4 }),
      ]);
      expect(fragment.borders.filter((border) => border.edge === 'bottom')).toEqual([
        expect.objectContaining({ widthPt: 4 }),
      ]);
      expect(fragment.inkBounds.yPt).toBeCloseTo(fragment.flowBounds.yPt - 2, 6);
      expect(fragment.inkBounds.heightPt).toBeCloseTo(fragment.flowBounds.heightPt + 4, 6);
    }
  });

  it('charges thick collapsed outer half-rules only to page-local non-exact tracks', () => {
    const outer = singleBorder(12);
    const source = table(
      [
        row([cell([], { marginTop: 5, marginBottom: 5 })], {
          rowHeight: null, rowHeightRule: 'auto',
        }),
        row([cell([])], { rowHeight: 1, rowHeightRule: 'exact' }),
        row([cell([], { marginTop: 5, marginBottom: 5 })], {
          rowHeight: null, rowHeightRule: 'auto',
        }),
        row([cell([])], { rowHeight: 1, rowHeightRule: 'exact' }),
        row([cell([], { marginTop: 5, marginBottom: 5 })], {
          rowHeight: null, rowHeightRule: 'auto',
        }),
      ],
      [120],
      {
        borders: {
          top: outer, bottom: outer, left: null, right: null,
          insideH: null, insideV: null,
        },
      },
    );

    expect(tableFormatInput(source).rows.map((format) => format.height?.rule ?? 'auto'))
      .toEqual(['auto', 'exact', 'auto', 'exact', 'auto']);

    const fragments = retainedTopLevelTables(
      layoutDocument(documentModel([source as unknown as BodyElement], 52)),
    );
    expect(fragments).toHaveLength(2);
    expect(fragments.map((fragment) => fragment.rows.map((row) => row.logicalRowIndex)))
      .toEqual([[0, 1, 2, 3], [4]]);
    expect(fragments.map((fragment) => fragment.rows.map((row) => row.heightPt)))
      .toEqual([[16, 1, 10, 1], [22]]);
    expect(fragments.map((fragment) => fragment.advancePt)).toEqual([28, 22]);
    expect(fragments.map((fragment) => fragment.inkBounds.heightPt)).toEqual([40, 34]);
  });

  it('re-resolves repeated atLeast header boundaries and row tracks on every slice', () => {
    const outer = singleBorder(4);
    const inside = singleBorder(1);
    const source = table(
      [
        row([textCell('header')], {
          isHeader: true, rowHeight: 20, rowHeightRule: 'atLeast',
        }),
        ...Array.from({ length: 3 }, (_unused, index) => row(
          [textCell(`body ${index}`)],
          { rowHeight: 20, rowHeightRule: 'atLeast' },
        )),
      ],
      [120],
      {
        borders: {
          top: outer, bottom: outer, left: null, right: null,
          insideH: inside, insideV: null,
        },
      },
    );

    const fragments = retainedTopLevelTables(
      layoutDocument(documentModel([source as unknown as BodyElement], 70)),
    );

    expect(fragments).toHaveLength(3);
    for (const [index, fragment] of fragments.entries()) {
      expect(fragment.rows.map((fragmentRow) => fragmentRow.heightPt)).toEqual([22.5, 22.5]);
      expect(fragment.advancePt).toBe(45);
      expect(fragment.rows[0]?.ownership).toBe(index === 0 ? 'source' : 'repeated-header');
      expect(fragment.borders.filter((border) => border.edge === 'top')).toEqual([
        expect.objectContaining({ widthPt: 4 }),
      ]);
      expect(fragment.borders.filter((border) => border.edge === 'between')).toEqual([
        expect.objectContaining({ widthPt: 1 }),
      ]);
      expect(fragment.borders.filter((border) => border.edge === 'bottom')).toEqual([
        expect.objectContaining({ widthPt: 4 }),
      ]);
    }
  });

  it('reserves the repeated header outer half-rule before admitting exact body rows', () => {
    const outer = singleBorder(4);
    const inside = singleBorder(1);
    const source = table(
      [
        row([textCell('header')], {
          isHeader: true, rowHeight: 20, rowHeightRule: 'atLeast',
        }),
        ...Array.from({ length: 4 }, (_unused, index) => row(
          [textCell(`body ${index}`)],
          { rowHeight: 20, rowHeightRule: 'exact' },
        )),
      ],
      [120],
      {
        borders: {
          top: outer, bottom: outer, left: null, right: null,
          insideH: inside, insideV: null,
        },
      },
    );

    const fragments = retainedTopLevelTables(
      layoutDocument(documentModel([source as unknown as BodyElement], 80)),
    );
    expect(fragments).toHaveLength(4);
    expect(fragments.map((fragment) => fragment.rows.length)).toEqual([2, 2, 2, 2]);
    expect(fragments.map((fragment) => fragment.advancePt)).toEqual([42.5, 42.5, 42.5, 42.5]);
    for (const [index, fragment] of fragments.entries()) {
      expect(fragment.rows.map((fragmentRow) => fragmentRow.logicalRowIndex)).toEqual([0, index + 1]);
      expect(fragment.borders.filter((border) => border.edge === 'between')).toHaveLength(1);
      expect(fragment.borders.filter((border) => border.edge === 'bottom')).toEqual([
        expect.objectContaining({ widthPt: 4 }),
      ]);
    }
  });

  it('retains split-row content advance and each piece own page-local boundary ink', () => {
    const outer = singleBorder(4);
    const inside = singleBorder(1);
    const source = table(
      [row([cell([{
        type: 'paragraph',
        ...paragraph('あ'.repeat(400)),
      } as unknown as CellElement])])],
      [120],
      {
        borders: {
          top: outer, bottom: outer, left: null, right: null,
          insideH: inside, insideV: null,
        },
      },
    );

    const fragments = retainedTopLevelTables(
      layoutDocument(documentModel([source as unknown as BodyElement], 80)),
    );

    expect(fragments.length).toBeGreaterThan(1);
    for (const [fragmentIndex, fragment] of fragments.entries()) {
      const fragmentRow = fragment.rows[0];
      expect(fragmentRow?.logicalRowIndex).toBe(0);
      expect(fragmentRow?.fragmentIndex).toBe(fragmentIndex);
      expect(fragment.advancePt).toBeCloseTo((fragmentRow?.contentHeightPt ?? 0) + 4, 8);
      expect(fragment.flowBounds.heightPt).toBeCloseTo(fragment.advancePt, 8);
      expect(fragment.borders.filter((border) => border.edge === 'top')).toEqual([
        expect.objectContaining({ widthPt: 4 }),
      ]);
      expect(fragment.borders.filter((border) => border.edge === 'bottom')).toEqual([
        expect.objectContaining({ widthPt: 4 }),
      ]);
      expect(fragment.inkBounds.yPt).toBeCloseTo(fragment.flowBounds.yPt - 2, 8);
      expect(fragment.inkBounds.heightPt).toBeCloseTo(fragment.advancePt + 4, 8);
    }
    const lineRanges = fragments.map((fragment) => {
      const range = fragment.rows[0]?.cells[0]?.contentRanges.find(
        (candidate) => candidate.kind === 'paragraph',
      );
      if (range?.kind !== 'paragraph') throw new Error('expected split paragraph ownership');
      return range;
    });
    expect(lineRanges[0]?.lineStart).toBe(0);
    for (let index = 1; index < lineRanges.length; index += 1) {
      expect(lineRanges[index]?.lineStart).toBe(lineRanges[index - 1]?.lineEnd);
    }
    expect(fragments.reduce((sum, fragment) => sum + fragment.advancePt, 0)).toBeCloseTo(
      fragments.reduce((sum, fragment) => sum + (fragment.rows[0]?.contentHeightPt ?? 0) + 4, 0),
      8,
    );
  });
});
