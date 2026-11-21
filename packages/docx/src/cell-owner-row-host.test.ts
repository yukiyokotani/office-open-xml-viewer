import { describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { textRunGeometryForPage } from './layout/text-index.js';
import { paintLayoutPage } from './paint/canvas-page.js';
import type { TableFragmentLayout } from './layout/table-pagination.js';
import type { DocumentLayout, LayoutRect, TableLayout } from './layout/types.js';
import type {
  BodyElement,
  BorderSpec,
  CellElement,
  DocNote,
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  FramePr,
  SectionProps,
  TableBorders,
} from './types.js';

// Cell-owner row hosts (table-owner-runs.ts): a row whose logical
// first cell starts with a non-drop-cap framePr paragraph, in a body, header
// or footer ROOT table whose effective §17.4.57 positioning is null, is
// placed as one owner at the §17.3.1.11 frame position. Tables nested in a
// cell and text box / note roots elect nothing (Word imports those frames as
// inert); cell-frame-owner-context.test.ts covers that eligibility through
// the real parser. The positioned-table cases below protect the page
// placement of §17.4.57 tables below table cells and in story tables, which
// the same pagination steps own. Synthetic linear metrics (one 10pt glyph
// advance per character, 10pt lines) are a test measurement service only.
// Expected geometry is derived from authored frame operands and from the same
// document laid out without a carrier, never read back from the
// implementation.

function measureContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText: (text: string) => ({
      width: [...text].length * px(),
      actualBoundingBoxAscent: px() * 0.8,
      actualBoundingBoxDescent: px() * 0.2,
      fontBoundingBoxAscent: px() * 0.8,
      fontBoundingBoxDescent: px() * 0.2,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

// Page 400×400pt with 40pt margins: body band x∈[40,360], y∈[40,360].
const MARGIN = 40;
const PAGE = 400;

function section(overrides: Partial<SectionProps> = {}): SectionProps {
  return {
    pageWidth: PAGE, pageHeight: PAGE,
    marginTop: MARGIN, marginRight: MARGIN, marginBottom: MARGIN, marginLeft: MARGIN,
    headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    ...overrides,
  } as SectionProps;
}

function paragraph(text: string, extra: Partial<DocParagraph> = {}): DocParagraph & { type: 'paragraph' } {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: text ? [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'NotInMetrics',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }] : [],
    defaultFontSize: 10, defaultFontFamily: 'NotInMetrics', widowControl: false,
    ...extra,
  } as unknown as DocParagraph & { type: 'paragraph' };
}

function frame(overrides: Partial<FramePr> = {}): FramePr {
  return {
    dropCap: 'none', lines: 1, wrap: 'notBeside', hAnchor: 'page', vAnchor: 'page',
    hRule: 'auto', hSpace: 0, vSpace: 0,
    ...overrides,
  };
}

function framed(text: string, framePr: FramePr): CellElement {
  return paragraph(text, { framePr }) as CellElement;
}

const single = (width = 1): BorderSpec => ({ style: 'single', width, color: '000000' });
const noBorders: TableBorders = {
  top: null, bottom: null, left: null, right: null, insideH: null, insideV: null,
};
// A cell's text sits below its row's top rule (WORD_COLLAPSED_BORDER_ROW_TRACK_FOOTPRINT).
const TOP_RULE_INSET = 1;
const boxBorders: TableBorders = {
  top: single(), bottom: single(), left: single(), right: single(),
  insideH: single(), insideV: single(),
};

function cell(content: CellElement[], extra: Partial<DocTableCell> = {}): DocTableCell {
  return {
    content, colSpan: 1, vMerge: null, borders: { ...noBorders },
    background: null, vAlign: 'top', widthPt: null,
    ...extra,
  };
}

function row(cells: DocTableCell[], extra: Partial<DocTableRow> = {}): DocTableRow {
  return { cells, rowHeight: null, rowHeightRule: 'auto', isHeader: false, ...extra };
}

/** One authored row: cell A holds three paragraphs (block 0 optionally the
 * carrier), cell B one ordinary marker. Labels are unique per row. */
function markerRow(
  label: string,
  carrier: Readonly<{ block: 0 | 1; framePr: FramePr }> | null,
  extra: Partial<DocTableRow> = {},
): DocTableRow {
  const a = [`${label}a1`, `${label}a2`, `${label}a3`].map((text, index) =>
    carrier?.block === index ? framed(text, carrier.framePr) : paragraph(text) as CellElement);
  return row([cell(a), cell([paragraph(`${label}b`) as CellElement])], extra);
}

function table(rows: DocTableRow[], extra: Partial<DocTable> = {}): BodyElement {
  return {
    type: 'table', colWidths: [100, 100], rows, borders: boxBorders,
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 5, cellMarginRight: 5,
    jc: 'left', layout: 'fixed',
    ...extra,
  } as unknown as BodyElement;
}

function layout(
  body: BodyElement[],
  options: Readonly<{
    section?: SectionProps; footnotes?: DocNote[]; endnotes?: DocNote[];
    header?: BodyElement[]; footer?: BodyElement[];
    allowFootnoteContinuation?: boolean;
  }> = {},
): DocumentLayout {
  const model = {
    section: options.section ?? section(),
    body,
    headers: {
      default: options.header ? { body: options.header } : null,
      first: null,
      even: null,
    },
    footers: {
      default: options.footer ? { body: options.footer } : null,
      first: null,
      even: null,
    },
    footnotes: options.footnotes ?? [], endnotes: options.endnotes ?? [], fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
  const before = JSON.stringify(model);
  const result = layoutDocument(
    model,
    createLayoutServices(model, {
      measureContext: measureContext(),
      ...(options.allowFootnoteContinuation ? { allowFootnoteContinuation: true } : {}),
    }),
    { currentDateMs: 0 },
  );
  // Source conservation: layout never rewrites the immutable source model,
  // so every framePr declaration remains retained.
  expect(JSON.stringify(model)).toBe(before);
  return result;
}

type Box = Readonly<{ text: string; x: number; y: number; path: readonly number[] }>;

function runs(result: DocumentLayout, pageIndex = 0): readonly Box[] {
  return textRunGeometryForPage(result, pageIndex).map((run) => ({
    text: run.placement.text,
    x: run.placement.bounds.xPt + run.pointToPage.e,
    y: run.placement.bounds.yPt + run.pointToPage.f,
    path: run.source.path,
  }));
}

function at(boxes: readonly Box[], text: string): Box {
  const matches = boxes.filter((box) => box.text === text);
  expect(matches, text).toHaveLength(1);
  return matches[0]!;
}

type Matrix = Readonly<{ a: number; b: number; c: number; d: number; e: number; f: number }>;

const IDENTITY: Matrix = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };

/** `m` after `n`, as Canvas 2D composes `transform(n)` onto the current `m`. */
const multiply = (m: Matrix, n: Matrix): Matrix => ({
  a: m.a * n.a + m.c * n.b,
  b: m.b * n.a + m.d * n.b,
  c: m.a * n.c + m.c * n.d,
  d: m.b * n.c + m.d * n.d,
  e: m.a * n.e + m.c * n.f + m.e,
  f: m.b * n.e + m.d * n.f + m.f,
});

/**
 * A Canvas 2D recorder for the production page painter: it keeps the full
 * current transform through save/restore, setTransform, transform, translate,
 * scale and rotate, and records every fillText at its device point and every
 * fillRect as its device-space bounding box with the fill style it was
 * painted in. It draws nothing else; any other method is a no-op. Painting
 * must not measure text, so measureText counts and throws.
 */
function recordingCanvas() {
  const texts: { text: string; x: number; y: number }[] = [];
  const fills: { color: string; x: number; y: number; width: number; height: number }[] = [];
  let measures = 0;
  let matrix = IDENTITY;
  const stack: Matrix[] = [];
  const state: Record<string | symbol, unknown> = {
    font: '10px serif', fillStyle: '#000000', strokeStyle: '#000000', lineWidth: 1,
    globalAlpha: 1, textAlign: 'left', textBaseline: 'alphabetic', direction: 'ltr',
    letterSpacing: '0px', fontKerning: 'auto', lineCap: 'butt', lineJoin: 'miter',
  };
  const canvas = { width: 0, height: 0, style: {}, getContext: () => ctx };
  const methods: Record<string, (...args: never[]) => unknown> = {
    save: () => { stack.push(matrix); },
    restore: () => { matrix = stack.pop() ?? matrix; },
    setTransform: (...args: (number | Matrix)[]) => {
      const [a, b, c, d, e, f] = args as number[];
      matrix = typeof args[0] === 'object' ? { ...(args[0] as Matrix) } : { a: a!, b: b!, c: c!, d: d!, e: e!, f: f! };
    },
    resetTransform: () => { matrix = IDENTITY; },
    getTransform: () => ({ ...matrix }),
    transform: (a: number, b: number, c: number, d: number, e: number, f: number) => {
      matrix = multiply(matrix, { a, b, c, d, e, f });
    },
    translate: (x: number, y: number) => { matrix = multiply(matrix, { a: 1, b: 0, c: 0, d: 1, e: x, f: y }); },
    scale: (x: number, y: number) => { matrix = multiply(matrix, { a: x, b: 0, c: 0, d: y, e: 0, f: 0 }); },
    rotate: (angle: number) => {
      const cos = Math.cos(angle);
      const sin = Math.sin(angle);
      matrix = multiply(matrix, { a: cos, b: sin, c: -sin, d: cos, e: 0, f: 0 });
    },
    measureText: () => {
      measures += 1;
      throw new Error('painting measured text');
    },
    fillText: (text: string, x: number, y: number) => {
      texts.push({ text, x: matrix.a * x + matrix.c * y + matrix.e, y: matrix.b * x + matrix.d * y + matrix.f });
    },
    fillRect: (x: number, y: number, width: number, height: number) => {
      const corners = [[x, y], [x + width, y], [x, y + height], [x + width, y + height]]
        .map(([px, py]) => [matrix.a * px! + matrix.c * py! + matrix.e, matrix.b * px! + matrix.d * py! + matrix.f]);
      const xs = corners.map(([px]) => px!);
      const ys = corners.map(([, py]) => py!);
      fills.push({
        color: String(state.fillStyle),
        x: Math.min(...xs), y: Math.min(...ys),
        width: Math.max(...xs) - Math.min(...xs), height: Math.max(...ys) - Math.min(...ys),
      });
    },
    createLinearGradient: () => ({ addColorStop() {} }),
    createRadialGradient: () => ({ addColorStop() {} }),
    createPattern: () => null,
  };
  const noop = () => undefined;
  const ctx: CanvasRenderingContext2D = new Proxy({} as CanvasRenderingContext2D, {
    get: (_, key) => {
      if (key === 'canvas') return canvas;
      if (typeof key === 'string' && key in methods) return methods[key];
      if (key in state) return state[key];
      return noop;
    },
    set: (_, key, value) => {
      state[key] = value;
      return true;
    },
    has: (_, key) => key === 'canvas' || (typeof key === 'string' && key in methods) || key in state,
  });
  return { canvas: canvas as unknown as HTMLCanvasElement, texts, fills, measures: () => measures };
}

function tables(result: DocumentLayout, pageIndex = 0): readonly TableFragmentLayout[] {
  return result.pages[pageIndex]!.layers.body
    .filter((node): node is TableLayout => node.kind === 'table') as TableFragmentLayout[];
}

const horizontalBorders = (node: TableLayout) => node.borders
  .filter((segment) => segment.from.yPt === segment.to.yPt)
  .map((segment) => ({ y: segment.from.yPt, style: segment.authoredStyle, widthPt: segment.widthPt }));

/** A one-column outer table whose cell holds PRE, `inner` (if any) and POST. */
const nestedIn = (inner: BodyElement | null, colWidths: number[]) => table([row([
  cell([
    paragraph('pre') as CellElement,
    ...(inner ? [inner as unknown as CellElement] : []),
    paragraph('post') as CellElement,
  ]),
])], { colWidths } as Partial<DocTable>);

describe('cell-owner row host at the production layout boundary', () => {
  // N1 / N2 / NF: identical one-row tables; N1 carries the frame on A-P1,
  // N2 on A-P2, NF none. Frame: page/page x=150 y=200, w = the 200pt grid.
  const hostFrame = frame({ x: 150, y: 200, w: 200 });
  const document = (carrier: Readonly<{ block: 0 | 1; framePr: FramePr }> | null) => layout([
    paragraph('before') as BodyElement,
    table([markerRow('r1', carrier)]),
    paragraph('after') as BodyElement,
  ]);

  it('moves N1 borders and ordinary siblings by one delta; N2 equals the carrier-free control', () => {
    const control = document(null);
    const n1 = document({ block: 0, framePr: hostFrame });
    const n2 = document({ block: 1, framePr: hostFrame });

    // N2: a later first-cell carrier elects nothing; geometry equals NF.
    expect(runs(n2)).toEqual(runs(control));
    expect(tables(n2).map((node) => [node.flowBounds, node.borders]))
      .toEqual(tables(control).map((node) => [node.flowBounds, node.borders]));

    const [controlTable] = tables(control);
    const [host] = tables(n1);
    expect(tables(n1)).toHaveLength(1);
    // The whole host, borders included, lands at the authored frame origin.
    expect(host!.flowBounds).toMatchObject({ xPt: 150, yPt: 200 });
    const dx = 150 - controlTable!.flowBounds.xPt;
    const dy = 200 - controlTable!.flowBounds.yPt;
    expect(host!.borders.map((segment) => [segment.from, segment.to, segment.authoredStyle]))
      .toEqual(controlTable!.borders.map((segment) => [
        { xPt: segment.from.xPt + dx, yPt: segment.from.yPt + dy },
        { xPt: segment.to.xPt + dx, yPt: segment.to.yPt + dy },
        segment.authoredStyle,
      ]));
    // Original row/cell source identity is retained by the host.
    expect(host!.rows.map((laidOut) => [laidOut.source, laidOut.cells.map((c) => c.source)]))
      .toEqual(controlTable!.rows.map((laidOut) => [laidOut.source, laidOut.cells.map((c) => c.source)]));

    const controlRuns = runs(control);
    const hostRuns = runs(n1);
    // Every cell paragraph (the carrier, its ordinary siblings and cell B)
    // moves by the same delta, and keeps its authored source.
    for (const text of ['r1a1', 'r1a2', 'r1a3', 'r1b']) {
      const before = at(controlRuns, text);
      expect(at(hostRuns, text)).toEqual({ ...before, x: before.x + dx, y: before.y + dy });
    }
    // Zero host advance: AFTER is admitted at the natural cursor the table
    // started at (the line after BEFORE), because the host's exclusion does
    // not intersect it.
    expect(at(hostRuns, 'before')).toEqual(at(controlRuns, 'before'));
    expect(at(hostRuns, 'after')).toMatchObject({
      x: at(controlRuns, 'after').x,
      y: at(hostRuns, 'before').y + 10,
    });
    // Reading order stays source order.
    expect(hostRuns.map((box) => box.text))
      .toEqual(['before', 'r1a1', 'r1a2', 'r1a3', 'r1b', 'after']);
  });

  it('places differing row tuples independently with table caps; equal tuples share one host', () => {
    const capped: TableBorders = {
      ...boxBorders,
      top: { style: 'double', width: 3, color: '000000' },
      bottom: { style: 'thick', width: 2, color: '000000' },
    };
    const twoRows = (first: FramePr, second: FramePr) => layout([
      paragraph('before') as BodyElement,
      table([
        markerRow('r1', { block: 0, framePr: first }),
        markerRow('r2', { block: 0, framePr: second }),
      ], { borders: capped }),
      paragraph('after') as BodyElement,
    ]);

    const unequal = twoRows(frame({ x: 150, y: 150 }), frame({ x: 60, y: 250 }));
    const [first, second] = tables(unequal);
    expect(tables(unequal)).toHaveLength(2);
    expect(first!.flowBounds).toMatchObject({ xPt: 150, yPt: 150 });
    expect(second!.flowBounds).toMatchObject({ xPt: 60, yPt: 250 });
    expect(first!.rows.map((laidOut) => laidOut.source.path)).toEqual([[1, 0]]);
    expect(second!.rows.map((laidOut) => laidOut.source.path)).toEqual([[1, 1]]);
    // A run edge is a segment edge: table top/bottom, never insideH.
    for (const host of [first!, second!]) {
      const edges = horizontalBorders(host);
      expect(edges.map((edge) => edge.style)).toEqual(['double', 'thick']);
      // The top and bottom rules lie inside the segment's box.
      expect(edges[0]!.y).toBeCloseTo(host.flowBounds.yPt + edges[0]!.widthPt / 2, 9);
      expect(edges[1]!.y).toBeCloseTo(host.flowBounds.yPt + host.flowBounds.heightPt - edges[1]!.widthPt / 2, 9);
    }
    const unequalRuns = runs(unequal);
    expect(unequalRuns.map((box) => box.text)).toEqual([
      'before', 'r1a1', 'r1a2', 'r1a3', 'r1b', 'r2a1', 'r2a2', 'r2a3', 'r2b', 'after',
    ]);
    // Both hosts lie below AFTER's line at the natural cursor.
    expect(at(unequalRuns, 'after').y).toBe(at(unequalRuns, 'before').y + 10);

    const equal = twoRows(frame({ x: 150, y: 150 }), frame({ x: 150, y: 150 }));
    const [shared] = tables(equal);
    expect(tables(equal)).toHaveLength(1);
    expect(shared!.flowBounds).toMatchObject({ xPt: 150, yPt: 150 });
    expect(shared!.rows.map((laidOut) => laidOut.source.path)).toEqual([[1, 0], [1, 1]]);
    expect(horizontalBorders(shared!).map((edge) => edge.style)).toEqual(['double', 'single', 'thick']);

    // hSpace is a tuple member: an otherwise equal pair forms two hosts.
    const spaced = twoRows(frame({ x: 150, y: 150, hSpace: 4 }), frame({ x: 150, y: 150 }));
    expect(tables(spaced)).toHaveLength(2);
  });

  describe('a vertical merge across rows whose carriers differ', () => {
    // Observed in Word 16.113.3 (macOS) on public synthetic sources: two rows,
    // a fixed two-column grid (2×189pt in the Word controls), second-column
    // restart/continue, page/page notBeside carriers on cell A's first
    // paragraph. Original-open and saved/reopened pages were exactly equal.
    // The merge does not make the rows share a placement: each row goes to its
    // own carrier's frame or stays in ordinary flow, and Word saves two one-row
    // tables (restart in the first, an empty continue in the second). Equal
    // carriers over a second-column merge were not probed natively: the
    // equal-carrier case below follows the grid model and ECMA-376 §17.4.84,
    // while native equal-carrier evidence covers unmerged rows and the
    // first-column merge (next describe). These tests reuse the authored frame
    // positions but use their own synthetic grid; every other expectation
    // derives from carrier-free controls, not from a measured Word scale.
    const first = frame({ x: 144, y: 90 });
    const later = frame({ x: 90, y: 210 });
    const lead = (text: string, carrier: FramePr | null) =>
      cell([carrier ? framed(text, carrier) : paragraph(text) as CellElement]);
    const mergedRows = (carrier0: FramePr | null, carrier1: FramePr | null) => [
      row([lead('m1a', carrier0), cell([paragraph('m1b') as CellElement], { vMerge: true })]),
      row([lead('m2a', carrier1), cell([paragraph('') as CellElement], { vMerge: false })]),
    ];
    // Carrier-free, merge-free one-row controls of each row.
    const plainRow0 = () => row([lead('m1a', null), cell([paragraph('m1b') as CellElement])]);
    const plainRow1 = () => row([lead('m2a', null), cell([paragraph('') as CellElement])]);
    const document = (rows: DocTableRow[]) => layout([
      paragraph('before') as BodyElement, table(rows), paragraph('after') as BodyElement,
    ]);
    const point = (box: Box) => ({ text: box.text, x: box.x, y: box.y });
    const paths = (node: TableLayout) => node.rows.map((laidOut) => laidOut.source.path);
    /** The single laid-out table that owns body table 1's source row `rowIndex`. */
    const ownerOf = (result: DocumentLayout, rowIndex: number) => {
      const owners = tables(result).filter((node) =>
        node.rows.some((laidOut) => laidOut.source.path.join() === `1,${rowIndex}`));
      expect(owners, `row ${rowIndex}`).toHaveLength(1);
      return owners[0]!;
    };
    /** `text` inside a host at `origin`, offset as in its one-row control. */
    const hosted = (control: DocumentLayout, text: string, origin: FramePr) => {
      const [controlTable] = tables(control);
      const box = at(runs(control), text);
      return {
        text,
        x: box.x - controlTable!.flowBounds.xPt + origin.x!,
        y: box.y - controlTable!.flowBounds.yPt + origin.y!,
      };
    };
    const carrierDiagnostics = (result: DocumentLayout) => result.diagnostics.filter((diagnostic) =>
      diagnostic.source?.story === 'body' && diagnostic.source.path.length === 4);
    const readingOrder = ['before', 'm1a', 'm1b', 'm2a', 'after'];

    it('places only the first row at its carrier; the later row stays in ordinary flow', () => {
      const result = document(mergedRows(first, null));
      const row1Control = document([plainRow1()]);
      expect(tables(result)).toHaveLength(2);
      const host = ownerOf(result, 0);
      expect(host).toMatchObject({ ordinaryFlow: false, flowBounds: { xPt: 144, yPt: 90 } });
      expect(paths(host)).toEqual([[1, 0]]);
      const ordinary = ownerOf(result, 1);
      expect(ordinary.ordinaryFlow).toBe(true);
      expect(paths(ordinary)).toEqual([[1, 1]]);
      // The host advances nothing, so the ordinary row lands where the
      // one-row control of that row does.
      expect(ordinary.flowBounds).toEqual(tables(row1Control)[0]!.flowBounds);
      const boxes = runs(result);
      for (const text of ['m1a', 'm1b']) {
        expect(point(at(boxes, text))).toEqual(hosted(document([plainRow0()]), text, first));
      }
      for (const text of ['before', 'm2a', 'after']) {
        expect(point(at(boxes, text))).toEqual(point(at(runs(row1Control), text)));
      }
      expect(boxes.map((box) => box.text)).toEqual(readingOrder);
    });

    it('paints the full grid of an ordinary row whose continuation lost its restart to a host', () => {
      // ECMA-376 §17.4.84 merges a continue cell with the restart above it in
      // the same grid column. The observed Word row-run partition above puts
      // the restart in the host and saves the later row as its own one-row
      // table with the continue retained; Word paints that row's second cell
      // as an empty bordered cell across the whole grid. So the orphan
      // continuation still owns its grid area (empty, with no restart text),
      // and the row's borders equal the merge-free control. A restart within
      // the same run keeps its merge: see the equal-carrier test below.
      // The host (y = 90) does not intersect this row's natural band, so
      // frame admission cannot move the row here.
      const result = document(mergedRows(first, null));
      const [controlTable] = tables(document([plainRow1()]));
      const ordinary = ownerOf(result, 1);
      const segments = (node: TableLayout) => node.borders
        .map((segment) => JSON.stringify([segment.from, segment.to, segment.authoredStyle]))
        .sort();
      // Outer rails of the control: top, bottom and right, with the extent
      // each one covers in the candidate.
      const ys = controlTable!.borders.flatMap((segment) => [segment.from.yPt, segment.to.yPt]);
      const xs = controlTable!.borders.flatMap((segment) => [segment.from.xPt, segment.to.xPt]);
      const top = Math.min(...ys);
      const bottom = Math.max(...ys);
      const right = Math.max(...xs);
      const extent = (node: TableLayout, onRail: (segment: TableLayout['borders'][number]) => boolean,
        axis: 'xPt' | 'yPt') => {
        const values = node.borders.filter(onRail)
          .flatMap((segment) => [segment.from[axis], segment.to[axis]]);
        return values.length ? [Math.min(...values), Math.max(...values)] : null;
      };
      const rails = (node: TableLayout) => ({
        top: extent(node, (s) => s.from.yPt === top && s.to.yPt === top, 'xPt'),
        bottom: extent(node, (s) => s.from.yPt === bottom && s.to.yPt === bottom, 'xPt'),
        right: extent(node, (s) => s.from.xPt === right && s.to.xPt === right, 'yPt'),
      });
      expect(rails(controlTable!).right).not.toBeNull();
      expect(rails(ordinary)).toEqual(rails(controlTable!));
      expect(segments(ordinary)).toEqual(segments(controlTable!));
      const boxes = runs(result);
      // Restart text paints once, in the host; the continuation paints nothing.
      expect(boxes.filter((box) => box.path.slice(0, 3).join() === '1,1,1')).toEqual([]);
      expect(boxes.map((box) => box.text)).toEqual(readingOrder);
    });

    /** Every laid-out table, on any page, that owns a source row at `rowPath`. */
    const fragmentsOf = (result: DocumentLayout, rowPath: string) => result.pages.flatMap((_, pageIndex) =>
      tables(result, pageIndex).filter((node) =>
        node.rows.some((laidOut) => laidOut.source.path.join() === rowPath)));
    const allRuns = (result: DocumentLayout) => result.pages.flatMap((_, pageIndex) => runs(result, pageIndex));

    it('paints the full grid in every fragment of a cut ordinary row whose continuation lost its restart', () => {
      // Source cut identity: a row cut across pages is the same source row in
      // every fragment, the first partial one included. So the orphan
      // continuation keeps its projected empty-owner grid area (library
      // projected-segment contract, test above) in each fragment, and each
      // fragment's borders equal those of the same ordinary row with a
      // merge-free empty second cell, cut alike under the same prefix.
      // PREFIX lines leave a 40pt band at the page foot; the eight-line first
      // cell cannot fit it, so the row is cut by real pagination. The host
      // (y = 90) lies above that band (precondition below).
      const PREFIX = 28;
      const prefix = Array.from({ length: PREFIX }, (_, index) => paragraph(`p${index}`) as BodyElement);
      const lines = Array.from({ length: 8 }, (_, index) => `k${index}`);
      const tall = () => cell(lines.map((text) => paragraph(text) as CellElement));
      const cutDocument = (rows: DocTableRow[]) => layout([
        ...prefix, table(rows), paragraph('after') as BodyElement,
      ]);
      const control = cutDocument([row([tall(), cell([paragraph('') as CellElement])])]);
      const result = cutDocument([
        row([lead('m1a', first), cell([paragraph('m1b') as CellElement], { vMerge: true })]),
        row([tall(), cell([paragraph('') as CellElement], { vMerge: false })]),
      ]);

      const controlFragments = fragmentsOf(control, `${PREFIX},0`);
      const ordinaryFragments = fragmentsOf(result, `${PREFIX},1`);
      // Precondition: the control row is really cut, into fragments 0..n-1.
      expect(controlFragments.length).toBeGreaterThan(1);
      expect(controlFragments.map((node) => node.rows.map((laidOut) => laidOut.fragmentIndex)))
        .toEqual(controlFragments.map((_, index) => [index]));
      // The candidate row is cut alike, each fragment an ordinary one-row table.
      expect(ordinaryFragments.map((node) => node.rows.map((laidOut) => laidOut.fragmentIndex)))
        .toEqual(controlFragments.map((_, index) => [index]));
      for (const node of ordinaryFragments) {
        expect(node.ordinaryFlow).toBe(true);
        expect(paths(node)).toEqual([[PREFIX, 1]]);
      }
      expect(ordinaryFragments.map((node) => node.flowBounds))
        .toEqual(controlFragments.map((node) => node.flowBounds));
      const hosts = fragmentsOf(result, `${PREFIX},0`);
      expect(hosts).toHaveLength(1);
      expect(hosts[0]).toMatchObject({ ordinaryFlow: false, flowBounds: { xPt: 144, yPt: 90 } });
      expect(hosts[0]!.flowBounds.yPt + hosts[0]!.flowBounds.heightPt)
        .toBeLessThanOrEqual(ordinaryFragments[0]!.flowBounds.yPt);

      const segments = (node: TableLayout) => node.borders
        .map((segment) => JSON.stringify([segment.from, segment.to, segment.authoredStyle]))
        .sort();
      const extent = (node: TableLayout, onRail: (segment: TableLayout['borders'][number]) => boolean,
        axis: 'xPt' | 'yPt') => {
        const values = node.borders.filter(onRail)
          .flatMap((segment) => [segment.from[axis], segment.to[axis]]);
        return values.length ? [Math.min(...values), Math.max(...values)] : null;
      };
      controlFragments.forEach((controlFragment, index) => {
        // Outer rails of this control fragment and the extent each covers.
        const ys = controlFragment.borders.flatMap((segment) => [segment.from.yPt, segment.to.yPt]);
        const xs = controlFragment.borders.flatMap((segment) => [segment.from.xPt, segment.to.xPt]);
        const top = Math.min(...ys);
        const bottom = Math.max(...ys);
        const left = Math.min(...xs);
        const right = Math.max(...xs);
        const rails = (node: TableLayout) => ({
          top: extent(node, (s) => s.from.yPt === top && s.to.yPt === top, 'xPt'),
          bottom: extent(node, (s) => s.from.yPt === bottom && s.to.yPt === bottom, 'xPt'),
          right: extent(node, (s) => s.from.xPt === right && s.to.xPt === right, 'yPt'),
        });
        // Precondition: the control fragment paints the whole authored grid.
        expect(rails(controlFragment), `control fragment ${index}`)
          .toEqual({ top: [left, right], bottom: [left, right], right: [top, bottom] });
        const candidate = ordinaryFragments[index]!;
        expect(rails(candidate), `fragment ${index}`).toEqual(rails(controlFragment));
        expect(segments(candidate), `fragment ${index}`).toEqual(segments(controlFragment));
      });

      const boxes = allRuns(result);
      // Restart text paints once, in the host; the continuation paints nothing.
      expect(boxes.filter((box) => box.path.slice(0, 3).join() === `${PREFIX},1,1`)).toEqual([]);
      expect(boxes.map((box) => box.text))
        .toEqual([...prefix.map((_, index) => `p${index}`), 'm1a', 'm1b', ...lines, 'after']);
    });

    it('charges a cut opening row only its own lines where the projected empty owner interval continues', () => {
      // ECMA-376 §17.4.68 tcMar and §17.4.80 trHeight under the library
      // interval policy (table.ts resolveRowTrack): the orphan continuation
      // opening the ordinary segment is a projected empty owner whose grid
      // interval runs through the continuations below it (rows 1..3). Its
      // authored margins are one interval constraint, whose deficit grows the
      // interval's last growable track, never an exact one. A cut fragment is
      // its own grid (library cut model): every cell of a cut row, the
      // projected owner included (source cut identity, test above), keeps
      // both margins in each fragment, and partialRow reserves them. So a
      // fragment holding only the opening row closes the owner there, at
      // TOP + BOTTOM; the completing fragment shares its grid with rows 2 and
      // 3, so there the opening row keeps exactly its own lines and the
      // deficit lands on row 3. No borders, so no rule or cap footprint
      // enters the tracks. LINE is the synthetic measurement line; the
      // suppressed continuations add none.
      //
      // Why this budget: a cut needs TOP + BOTTOM + LINE of band (the
      // reserve), and in such a band the uncut opening row (its track, or
      // the owner's requirement where that is larger) fits unless its own
      // lines exceed the band. So a cut opening row is longer than the first
      // band (29 lines > 240pt), and uncut its interval has no deficit; the
      // deficit arises only in the completing fragment. On a 360pt page the
      // table starts at y = 80: 6 lines fit beside the margins in the 240pt
      // first band, 10 in each 280pt fresh band, and the last 3 complete the
      // row on page 4 with rows 2 and 3 (30 + 20 + 110 < 180). Charging that
      // deficit to the completing row instead (a track window closed before
      // row 3) would leave 280 − 160 − 20 = 100pt, less than row 3's 110, and
      // cut row 3.
      const TOP = 90;
      const BOTTOM = 90;
      const LINE = 10;
      const CUT_PAGE = 360;
      const PREFIX = 4;
      const prefix = Array.from({ length: PREFIX }, (_, index) => paragraph(`p${index}`) as BodyElement);
      const above = frame({ x: 144, y: 40, w: 200 });
      const opening = Array.from({ length: 29 }, (_, index) => `q${index}`);
      const later = Array.from({ length: 11 }, (_, index) => `y${index}`);
      const lines = (texts: readonly string[]) => cell(texts.map((text) => paragraph(text) as CellElement));
      const orphan = (extra: Partial<DocTableCell> = {}) =>
        cell([paragraph('') as CellElement], { vMerge: false, ...extra });
      const rows = [
        row([lead('m1a', above), cell([paragraph('m1b') as CellElement], { vMerge: true })]),
        row([lines(opening), orphan({ marginTop: TOP, marginBottom: BOTTOM })]),
        row([cell([paragraph('x2') as CellElement]), orphan()], { rowHeight: 20, rowHeightRule: 'exact' }),
        row([lines(later), orphan()]),
      ];
      const intervalDocument = (pageHeight: number) => layout([
        ...prefix, table(rows, { borders: noBorders }), paragraph('after') as BodyElement,
      ], { section: section({ pageHeight }) });
      /** Source row index -> allocated height of each of its fragments, in page order. */
      const allocation = (result: DocumentLayout) => new Map([1, 2, 3].map((rowIndex) => [
        rowIndex,
        fragmentsOf(result, `${PREFIX},${rowIndex}`).flatMap((node) => {
          expect(node.ordinaryFlow).toBe(true);
          return node.rows
            .filter((laidOut) => laidOut.ownership === 'source'
              && laidOut.source.path.join() === `${PREFIX},${rowIndex}`)
            .map((laidOut) => laidOut.heightPt);
        }),
      ]));
      const expectMarkersOnce = (result: DocumentLayout) => {
        const boxes = allRuns(result);
        expect(boxes.map((box) => box.text)).toEqual([
          ...prefix.map((_, index) => `p${index}`), 'm1a', 'm1b', ...opening, 'x2', ...later, 'after',
        ]);
        for (const rowIndex of [1, 2, 3]) {
          expect(boxes.filter((box) => box.path.slice(0, 3).join() === `${PREFIX},${rowIndex},1`)).toEqual([]);
        }
      };

      // Whole control: 800pt page, nothing cut. Preconditions from authored
      // inputs: each row is its paragraph lines or its trHeight, since the
      // opening row alone exceeds TOP + BOTTOM.
      const whole = intervalDocument(800);
      const wholeRows = allocation(whole);
      expect(wholeRows.get(1), 'whole opening row').toEqual([opening.length * LINE]);
      expect(wholeRows.get(2), 'whole exact row').toEqual([20]);
      expect(wholeRows.get(3), 'whole later row').toEqual([later.length * LINE]);
      expect(fragmentsOf(whole, `${PREFIX},0`)[0]).toMatchObject({ ordinaryFlow: false, flowBounds: { xPt: 144, yPt: 40 } });
      expectMarkersOnce(whole);

      // Candidate: 360pt page, body bottom 320, so the opening row is cut.
      const cut = intervalDocument(CUT_PAGE);
      const hosts = fragmentsOf(cut, `${PREFIX},0`);
      expect(hosts).toHaveLength(1);
      expect(hosts[0]).toMatchObject({ ordinaryFlow: false, flowBounds: { xPt: 144, yPt: 40 } });
      const openingFragments = fragmentsOf(cut, `${PREFIX},1`);
      // Preconditions: the host lies above the ordinary band, and the opening
      // row is really cut, starting on the first page at the prefix's end.
      expect(openingFragments.length).toBeGreaterThan(1);
      expect(tables(cut, 0)).toContain(openingFragments[0]);
      expect(openingFragments[0]!.flowBounds.yPt).toBe(MARGIN + PREFIX * LINE);
      expect(hosts[0]!.flowBounds.yPt + hosts[0]!.flowBounds.heightPt)
        .toBeLessThanOrEqual(openingFragments[0]!.flowBounds.yPt);
      // The completing fragment holds the rest of the owner's interval.
      expect(openingFragments.at(-1)!.rows.map((laidOut) => laidOut.source.path.join()))
        .toEqual([1, 2, 3].map((rowIndex) => `${PREFIX},${rowIndex}`));
      const cutRows = allocation(cut);
      const completionPt = 3 * LINE;
      expect(cutRows.get(1), 'opening row').toEqual([TOP + BOTTOM, TOP + BOTTOM, TOP + BOTTOM, completionPt]);
      expect(cutRows.get(2), 'exact row').toEqual([20]);
      expect(cutRows.get(3), 'later growable row').toEqual([TOP + BOTTOM - completionPt - 20]);
      expectMarkersOnce(cut);
    });

    it('places conflicting carriers each at its own frame, neither suppressed', () => {
      const result = document(mergedRows(first, later));
      expect(tables(result)).toHaveLength(2);
      const firstHost = ownerOf(result, 0);
      const laterHost = ownerOf(result, 1);
      expect(firstHost).toMatchObject({ ordinaryFlow: false, flowBounds: { xPt: 144, yPt: 90 } });
      expect(laterHost).toMatchObject({ ordinaryFlow: false, flowBounds: { xPt: 90, yPt: 210 } });
      expect(paths(firstHost)).toEqual([[1, 0]]);
      expect(paths(laterHost)).toEqual([[1, 1]]);
      // Both carriers are applied; none is reported as unapplied.
      expect(carrierDiagnostics(result)).toEqual([]);
      const boxes = runs(result);
      for (const text of ['m1a', 'm1b']) {
        expect(point(at(boxes, text))).toEqual(hosted(document([plainRow0()]), text, first));
      }
      expect(point(at(boxes, 'm2a'))).toEqual(hosted(document([plainRow1()]), 'm2a', later));
      // Neither host advances the flow.
      expect(at(boxes, 'after')).toMatchObject({ x: MARGIN, y: at(boxes, 'before').y + 10 });
      expect(boxes.map((box) => box.text)).toEqual(readingOrder);
    });

    it('keeps the first row in ordinary flow when only the later row carries a frame', () => {
      const result = document(mergedRows(null, later));
      const row0Control = document([plainRow0()]);
      expect(tables(result)).toHaveLength(2);
      const ordinary = ownerOf(result, 0);
      expect(ordinary.ordinaryFlow).toBe(true);
      expect(paths(ordinary)).toEqual([[1, 0]]);
      expect(ordinary.flowBounds).toEqual(tables(row0Control)[0]!.flowBounds);
      const host = ownerOf(result, 1);
      expect(host).toMatchObject({ ordinaryFlow: false, flowBounds: { xPt: 90, yPt: 210 } });
      expect(paths(host)).toEqual([[1, 1]]);
      const boxes = runs(result);
      for (const text of ['before', 'm1a', 'm1b', 'after']) {
        expect(point(at(boxes, text))).toEqual(point(at(runs(row0Control), text)));
      }
      expect(point(at(boxes, 'm2a'))).toEqual(hosted(document([plainRow1()]), 'm2a', later));
      expect(boxes.map((box) => box.text)).toEqual(readingOrder);
    });

    it('keeps equal carriers one host whose merged grid translates the carrier-free table', () => {
      const control = document(mergedRows(null, null));
      const result = document(mergedRows(first, first));
      expect(tables(result)).toHaveLength(1);
      const [host] = tables(result);
      const [controlTable] = tables(control);
      expect(host!.flowBounds).toMatchObject({ xPt: 144, yPt: 90 });
      expect(paths(host!)).toEqual([[1, 0], [1, 1]]);
      const dx = 144 - controlTable!.flowBounds.xPt;
      const dy = 90 - controlTable!.flowBounds.yPt;
      // Borders, with no insideH under the merged cell, move by one delta.
      expect(host!.borders.map((segment) => [segment.from, segment.to, segment.authoredStyle]))
        .toEqual(controlTable!.borders.map((segment) => [
          { xPt: segment.from.xPt + dx, yPt: segment.from.yPt + dy },
          { xPt: segment.to.xPt + dx, yPt: segment.to.yPt + dy },
          segment.authoredStyle,
        ]));
      const controlRuns = runs(control);
      const boxes = runs(result);
      for (const text of ['m1a', 'm1b', 'm2a']) {
        const before = at(controlRuns, text);
        expect(at(boxes, text)).toEqual({ ...before, x: before.x + dx, y: before.y + dy });
      }
      expect(boxes.map((box) => box.text)).toEqual(readingOrder);
    });

    it('places conflicting header-story carriers each at its own frame', () => {
      // The header root takes the whole-table story path, not body
      // pagination. Extending the body observation to it is an inference.
      const bands = section({ headerDistance: 20 });
      const story = (rows: DocTableRow[]) => runs(layout([paragraph('body') as BodyElement], {
        section: bands,
        header: [table(rows), paragraph('hdr') as BodyElement],
      }));
      const control = story([plainRow0()]);
      // The carrier-free table starts the header band at (MARGIN, headerDistance).
      const inside = (text: string) => ({ x: at(control, text).x - MARGIN, y: at(control, text).y - 20 });
      const boxes = story(mergedRows(
        frame({ x: 144, y: 90, wrap: 'none' }),
        frame({ x: 90, y: 210, wrap: 'none' }),
      ));
      for (const text of ['m1a', 'm1b']) {
        expect(at(boxes, text)).toMatchObject({ x: 144 + inside(text).x, y: 90 + inside(text).y });
      }
      expect(at(boxes, 'm2a')).toMatchObject({ x: 90 + inside('m1a').x, y: 210 + inside('m1a').y });
      // Neither host advances the story.
      expect(at(boxes, 'hdr')).toMatchObject({ x: MARGIN, y: at(control, 'm1a').y - TOP_RULE_INSET });
    });
  });

  describe('a first-column vertical merge whose continuation row states its own frame', () => {
    // Observed in Word 16.113.3 (macOS) on public synthetic sources: two rows,
    // a fixed two-column grid (2×189pt in the Word controls), no explicit row
    // heights, first-column restart/continue. The restart cell holds three
    // marker paragraphs, the first a page/page notBeside carrier at 144/90
    // with w = the 378pt grid; the continuation cell holds one empty paragraph
    // whose pPr may state its own frame; the second column of the later row
    // holds a visible marker. Original-open and saved/reopened pages were
    // exactly equal, and the saved restart carrier was retained in each case:
    // - an equal continuation frame kept one two-row table, the later marker
    //   at its natural grid offset;
    // - a different continuation frame (90/210) saved two one-row tables,
    //   each row at its own frame;
    // - no continuation frame saved two one-row tables, the later row in
    //   ordinary flow (the restart carrier is not inherited).
    // The continuation's empty paragraph paints nothing, but its own frame
    // takes part in row identity. These tests reuse the authored frame
    // positions with w = their own synthetic 200pt grid; every other
    // expectation derives from carrier-free controls. The selector is shared
    // by the body and header/footer story roots (one acquisition call site),
    // so the body path stands for both here.
    const first = frame({ x: 144, y: 90, w: 200 });
    const later = frame({ x: 90, y: 210, w: 200 });
    const restart = (carrier: FramePr | null) => cell(['c1a1', 'c1a2', 'c1a3'].map((text, index) =>
      index === 0 && carrier ? framed(text, carrier) : paragraph(text) as CellElement), { vMerge: true });
    const continuation = (carrier: FramePr | null) =>
      cell([paragraph('', carrier ? { framePr: carrier } : {}) as CellElement], { vMerge: false });
    const mergedRows = (carrier0: FramePr | null, carrier1: FramePr | null) => [
      row([restart(carrier0), cell([paragraph('c1b') as CellElement])]),
      row([continuation(carrier1), cell([paragraph('c2b') as CellElement])]),
    ];
    // Carrier-free, merge-free one-row controls of each row.
    const plainRow0 = () => row([
      cell(['c1a1', 'c1a2', 'c1a3'].map((text) => paragraph(text) as CellElement)),
      cell([paragraph('c1b') as CellElement]),
    ]);
    const plainRow1 = () => row([cell([paragraph('') as CellElement]), cell([paragraph('c2b') as CellElement])]);
    const document = (rows: DocTableRow[]) => layout([
      paragraph('before') as BodyElement, table(rows), paragraph('after') as BodyElement,
    ]);
    const point = (box: Box) => ({ text: box.text, x: box.x, y: box.y });
    const paths = (node: TableLayout) => node.rows.map((laidOut) => laidOut.source.path);
    const ownerOf = (result: DocumentLayout, rowIndex: number) => {
      const owners = tables(result).filter((node) =>
        node.rows.some((laidOut) => laidOut.source.path.join() === `1,${rowIndex}`));
      expect(owners, `row ${rowIndex}`).toHaveLength(1);
      return owners[0]!;
    };
    const hosted = (control: DocumentLayout, text: string, origin: FramePr) => {
      const [controlTable] = tables(control);
      const box = at(runs(control), text);
      return {
        text,
        x: box.x - controlTable!.flowBounds.xPt + origin.x!,
        y: box.y - controlTable!.flowBounds.yPt + origin.y!,
      };
    };
    const carrierDiagnostics = (result: DocumentLayout) => result.diagnostics.filter((diagnostic) =>
      diagnostic.source?.story === 'body' && diagnostic.source.path.length === 4);
    // Visible markers once each in source order; the continuation paragraph
    // (body table 1, row 1, cell 0) paints nothing.
    const expectMarkers = (boxes: readonly Box[]) => {
      expect(boxes.map((box) => box.text))
        .toEqual(['before', 'c1a1', 'c1a2', 'c1a3', 'c1b', 'c2b', 'after']);
      expect(boxes.filter((box) => box.path.slice(0, 3).join() === '1,1,0')).toEqual([]);
    };

    it('keeps an equal continuation frame one host whose merged grid translates the carrier-free table', () => {
      const control = document(mergedRows(null, null));
      const result = document(mergedRows(first, first));
      expect(tables(result)).toHaveLength(1);
      const [host] = tables(result);
      const [controlTable] = tables(control);
      expect(host).toMatchObject({ ordinaryFlow: false, flowBounds: { xPt: 144, yPt: 90 } });
      expect(paths(host!)).toEqual([[1, 0], [1, 1]]);
      const dx = 144 - controlTable!.flowBounds.xPt;
      const dy = 90 - controlTable!.flowBounds.yPt;
      // Borders, with no insideH under the merged first cell, move by one delta.
      expect(host!.borders.map((segment) => [segment.from, segment.to, segment.authoredStyle]))
        .toEqual(controlTable!.borders.map((segment) => [
          { xPt: segment.from.xPt + dx, yPt: segment.from.yPt + dy },
          { xPt: segment.to.xPt + dx, yPt: segment.to.yPt + dy },
          segment.authoredStyle,
        ]));
      const controlRuns = runs(control);
      const boxes = runs(result);
      // The later row's marker stays at its natural merged-grid offset.
      for (const text of ['c1a1', 'c1a2', 'c1a3', 'c1b', 'c2b']) {
        const before = at(controlRuns, text);
        expect(at(boxes, text)).toEqual({ ...before, x: before.x + dx, y: before.y + dy });
      }
      // Zero host advance.
      expect(at(boxes, 'after')).toMatchObject({ x: MARGIN, y: at(boxes, 'before').y + 10 });
      expectMarkers(boxes);
    });

    it('places a different continuation frame as its own host; the later second cell moves with it', () => {
      const result = document(mergedRows(first, later));
      expect(tables(result)).toHaveLength(2);
      const firstHost = ownerOf(result, 0);
      const laterHost = ownerOf(result, 1);
      expect(firstHost).toMatchObject({ ordinaryFlow: false, flowBounds: { xPt: 144, yPt: 90 } });
      expect(laterHost).toMatchObject({ ordinaryFlow: false, flowBounds: { xPt: 90, yPt: 210 } });
      expect(paths(firstHost)).toEqual([[1, 0]]);
      expect(paths(laterHost)).toEqual([[1, 1]]);
      // Both carriers are applied; none is reported as unapplied.
      expect(carrierDiagnostics(result)).toEqual([]);
      const boxes = runs(result);
      for (const text of ['c1a1', 'c1a2', 'c1a3', 'c1b']) {
        expect(point(at(boxes, text))).toEqual(hosted(document([plainRow0()]), text, first));
      }
      expect(point(at(boxes, 'c2b'))).toEqual(hosted(document([plainRow1()]), 'c2b', later));
      // Neither host advances the flow.
      expect(at(boxes, 'after')).toMatchObject({ x: MARGIN, y: at(boxes, 'before').y + 10 });
      expectMarkers(boxes);
    });

    it('keeps a frame-free continuation row in ordinary flow without inheriting the restart carrier', () => {
      const result = document(mergedRows(first, null));
      const row1Control = document([plainRow1()]);
      expect(tables(result)).toHaveLength(2);
      const host = ownerOf(result, 0);
      expect(host).toMatchObject({ ordinaryFlow: false, flowBounds: { xPt: 144, yPt: 90 } });
      expect(paths(host)).toEqual([[1, 0]]);
      const ordinary = ownerOf(result, 1);
      expect(ordinary.ordinaryFlow).toBe(true);
      expect(paths(ordinary)).toEqual([[1, 1]]);
      // The host advances nothing, so the ordinary row lands where the
      // one-row control of that row does.
      expect(ordinary.flowBounds).toEqual(tables(row1Control)[0]!.flowBounds);
      const boxes = runs(result);
      for (const text of ['c1a1', 'c1a2', 'c1a3', 'c1b']) {
        expect(point(at(boxes, text))).toEqual(hosted(document([plainRow0()]), text, first));
      }
      for (const text of ['before', 'c2b', 'after']) {
        expect(point(at(boxes, text))).toEqual(point(at(runs(row1Control), text)));
      }
      expectMarkers(boxes);
    });

    // The host above lies below the ordinary row's natural band. Here the
    // carrier's authored y straddles the natural cursor (the line after
    // BEFORE, MARGIN + 10), so the ordinary row's prospective band
    // intersects the host's exclusion.
    const natural = MARGIN + 10;

    it('admits a frame-free continuation row below an intersecting notBeside host exclusion', () => {
      // §17.18.104 notBeside: no flow content beside the frame. The ordinary
      // row keeps its own x and is admitted at the host's bottom.
      const covering = frame({ x: 144, y: natural - 5, w: 200 });
      const row0Control = tables(document([plainRow0()]))[0]!;
      const row1Control = document([plainRow1()]);
      const [row1Table] = tables(row1Control);
      const hostBottomPt = covering.y! + row0Control.flowBounds.heightPt;
      // Precondition: the ordinary row's natural band intersects the host,
      // horizontally and vertically.
      expect(row1Table!.flowBounds.xPt + row1Table!.flowBounds.widthPt).toBeGreaterThan(covering.x!);
      expect(row1Table!.flowBounds.yPt).toBeLessThan(hostBottomPt);
      expect(row1Table!.flowBounds.yPt + row1Table!.flowBounds.heightPt).toBeGreaterThan(covering.y!);

      const result = document(mergedRows(covering, null));
      expect(tables(result)).toHaveLength(2);
      const host = ownerOf(result, 0);
      expect(host).toMatchObject({
        ordinaryFlow: false,
        flowBounds: { xPt: 144, yPt: covering.y!, heightPt: row0Control.flowBounds.heightPt },
      });
      expect(paths(host)).toEqual([[1, 0]]);
      const ordinary = ownerOf(result, 1);
      expect(ordinary.ordinaryFlow).toBe(true);
      expect(paths(ordinary)).toEqual([[1, 1]]);
      const dy = hostBottomPt - natural;
      expect(ordinary.flowBounds).toEqual({ ...row1Table!.flowBounds, yPt: row1Table!.flowBounds.yPt + dy });
      const boxes = runs(result);
      for (const text of ['c1a1', 'c1a2', 'c1a3', 'c1b']) {
        expect(point(at(boxes, text))).toEqual(hosted(document([plainRow0()]), text, covering));
      }
      expect(point(at(boxes, 'before'))).toEqual(point(at(runs(row1Control), 'before')));
      for (const text of ['c2b', 'after']) {
        const before = at(runs(row1Control), text);
        expect(point(at(boxes, text))).toEqual({ text, x: before.x, y: before.y + dy });
      }
      expectMarkers(boxes);
    });

    it('keeps a frame-free continuation row in place beside a horizontally disjoint around host', () => {
      // An around exclusion (w + hSpace) clear of the ordinary row's band
      // does not move it: a frame is not a full-width blocker.
      const beside = frame({ x: 250, y: natural - 5, w: 100, wrap: 'around' });
      const row1Control = document([plainRow1()]);
      const [row1Table] = tables(row1Control);
      // Precondition: vertically intersecting, horizontally disjoint.
      expect(row1Table!.flowBounds.xPt + row1Table!.flowBounds.widthPt).toBeLessThanOrEqual(beside.x!);
      expect(row1Table!.flowBounds.yPt + row1Table!.flowBounds.heightPt).toBeGreaterThan(beside.y!);

      const result = document(mergedRows(beside, null));
      expect(ownerOf(result, 0)).toMatchObject({ ordinaryFlow: false, flowBounds: { xPt: 250, yPt: beside.y! } });
      const ordinary = ownerOf(result, 1);
      expect(ordinary.ordinaryFlow).toBe(true);
      expect(ordinary.flowBounds).toEqual(row1Table!.flowBounds);
      const boxes = runs(result);
      for (const text of ['before', 'c2b', 'after']) {
        expect(point(at(boxes, text))).toEqual(point(at(runs(row1Control), text)));
      }
      expectMarkers(boxes);
    });

    it('admits a frame-free header continuation row below an intersecting notBeside host exclusion', () => {
      // §17.18.104 notBeside: no flow content beside the frame. This is the
      // common frame wrap rule applied to the header root's whole-story table
      // path; the native first-column observation above is body-only, so no
      // header pair is claimed here. HPRE puts the ordinary cursor one line
      // below the header band top; the page/page host straddles that cursor.
      const bands = section({ headerDistance: 20 });
      const story = (rows: DocTableRow[]) => layout([paragraph('body') as BodyElement], {
        section: bands,
        header: [paragraph('hpre') as BodyElement, table(rows), paragraph('hdr') as BodyElement],
      });
      const headerTables = (result: DocumentLayout) => result.pages[0]!.layers.header
        .filter((node): node is TableLayout => node.kind === 'table');
      /** The single header table that owns header table 1's source row `rowIndex`. */
      const headerOwnerOf = (result: DocumentLayout, rowIndex: number) => {
        const owners = headerTables(result).filter((node) =>
          node.rows.some((laidOut) => laidOut.source.path.join() === `1,${rowIndex}`));
        expect(owners, `header row ${rowIndex}`).toHaveLength(1);
        return owners[0]!;
      };
      const row0Control = story([plainRow0()]);
      const row1Control = story([plainRow1()]);
      const [row0Table] = headerTables(row0Control);
      const [row1Table] = headerTables(row1Control);
      // Retained story geometry is translated onto the page by the band: take
      // that translation from HPRE's retained node and its indexed page run.
      const hpreNode = row1Control.pages[0]!.layers.header.find((node) =>
        node.kind === 'paragraph' && node.source.path.join() === '0')!;
      const hpre = at(runs(row1Control), 'hpre');
      const toPage = { x: hpre.x - hpreNode.flowBounds.xPt, y: hpre.y - hpreNode.flowBounds.yPt };
      // The ordinary table starts on the line after HPRE.
      const headerNatural = hpre.y + 10;
      expect(row1Table!.flowBounds.yPt + toPage.y).toBe(headerNatural);
      const covering = frame({ x: 144, y: headerNatural - 5, w: 200 });
      // Host extent: the carrier-free restart row's own height.
      const hostBottomPt = covering.y! + row0Table!.flowBounds.heightPt;
      // Precondition: the ordinary row's natural page band intersects the
      // host, horizontally and vertically.
      const naturalX = row1Table!.flowBounds.xPt + toPage.x;
      expect(naturalX + row1Table!.flowBounds.widthPt).toBeGreaterThan(covering.x!);
      expect(naturalX).toBeLessThan(covering.x! + covering.w!);
      expect(headerNatural).toBeLessThan(hostBottomPt);
      expect(headerNatural + row1Table!.flowBounds.heightPt).toBeGreaterThan(covering.y!);

      const result = story(mergedRows(covering, null));
      const boxes = runs(result);
      // The host's markers sit at the frame origin, offset as in the
      // restart row's control.
      const row0Runs = runs(row0Control);
      for (const text of ['c1a1', 'c1a2', 'c1a3', 'c1b']) {
        const box = at(row0Runs, text);
        expect(point(at(boxes, text))).toEqual({
          text,
          x: box.x - (row0Table!.flowBounds.xPt + toPage.x) + covering.x!,
          y: box.y - (row0Table!.flowBounds.yPt + toPage.y) + covering.y!,
        });
      }
      const ordinary = headerOwnerOf(result, 1);
      expect(ordinary.ordinaryFlow).toBe(true);
      expect(paths(ordinary)).toEqual([[1, 1]]);
      // Admitted at the host's bottom, keeping its own x and width.
      const dy = hostBottomPt - headerNatural;
      expect(ordinary.flowBounds).toEqual({ ...row1Table!.flowBounds, yPt: row1Table!.flowBounds.yPt + dy });
      const row1Runs = runs(row1Control);
      expect(point(at(boxes, 'hpre'))).toEqual(point(at(row1Runs, 'hpre')));
      for (const text of ['c2b', 'hdr']) {
        const before = at(row1Runs, text);
        expect(point(at(boxes, text))).toEqual({ text, x: before.x, y: before.y + dy });
      }
      // Header markers once each in source order; the continuation paragraph
      // (header table 1, row 1, cell 0) paints nothing.
      const header = boxes.filter((box) => box.text !== 'body');
      expect(header.map((box) => box.text))
        .toEqual(['hpre', 'c1a1', 'c1a2', 'c1a3', 'c1b', 'c2b', 'hdr']);
      expect(header.filter((box) => box.path.slice(0, 3).join() === '1,1,0')).toEqual([]);
    });

    it('paints the authored shading of a header orphan continuation over its full cell', async () => {
      // Library contract for the projected segment, not a native header pair:
      // an orphan continuation in an ordinary row still owns its grid area
      // (see the body border test above), and §17.4.32 tcPr/shd shades that
      // cell's extent whether or not it holds text. So the header's ordinary
      // row paints the continuation's authored
      // fill exactly as a merge-free one-row control paints the same fill on
      // an ordinary empty cell in the same grid and band. The page/page host
      // lies far below the header's natural band, so admission cannot move
      // the ordinary row and only the shading is compared.
      const SHADE = 'C0FFEE';
      const away = frame({ x: 144, y: 210, w: 200 });
      const bands = section({ headerDistance: 20 });
      const story = (rows: DocTableRow[]) => layout([paragraph('body') as BodyElement], {
        section: bands,
        header: [table(rows), paragraph('hdr') as BodyElement],
      });
      const headerTables = (result: DocumentLayout) => result.pages[0]!.layers.header
        .filter((node): node is TableLayout => node.kind === 'table');
      const paint = async (result: DocumentLayout) => {
        const recorder = recordingCanvas();
        await paintLayoutPage(result, 0, recorder.canvas, { dpr: 1, scale: 1 }, { paint() {} });
        expect(recorder.measures()).toBe(0);
        return {
          shaded: recorder.fills.filter((fill) => fill.color.replace(/^#/, '').toUpperCase() === SHADE),
          texts: recorder.texts.map((call) => call.text).sort(),
        };
      };

      const control = story([row([
        cell([paragraph('') as CellElement], { background: SHADE }),
        cell([paragraph('c2b') as CellElement]),
      ])]);
      const result = story([
        row([restart(away), cell([paragraph('c1b') as CellElement])]),
        row([
          { ...continuation(null), background: SHADE },
          cell([paragraph('c2b') as CellElement]),
        ]),
      ]);

      // The ordinary row is placed exactly as the control table.
      const [controlTable] = headerTables(control);
      const ordinary = headerTables(result).filter((node) =>
        node.rows.some((laidOut) => laidOut.source.path.join() === '0,1'));
      expect(ordinary).toHaveLength(1);
      expect(ordinary[0]!.ordinaryFlow).toBe(true);
      expect(paths(ordinary[0]!)).toEqual([[0, 1]]);
      expect(ordinary[0]!.flowBounds).toEqual(controlTable!.flowBounds);

      const controlPaint = await paint(control);
      // Precondition: the control paints the authored fill once, as a real
      // region, and that region lies wholly above the host.
      expect(controlPaint.shaded).toHaveLength(1);
      const [fill] = controlPaint.shaded;
      expect(fill!.width).toBeGreaterThan(0);
      expect(fill!.height).toBeGreaterThan(0);
      expect(fill!.y + fill!.height).toBeLessThan(away.y!);

      const candidate = await paint(result);
      expect(candidate.shaded).toEqual(controlPaint.shaded);
      // The restart text paints once, in the host; the continuation paints none.
      expect(candidate.texts).toEqual(['body', 'c1a1', 'c1a2', 'c1a3', 'c1b', 'c2b', 'hdr']);
      expect(runs(result).filter((box) => box.path.slice(0, 3).join() === '0,1,0')).toEqual([]);
    });

    it('grows an opening orphan auto row to contain its continuation cell margins', () => {
      // ECMA-376 §17.4.68 tcMar: the continuation's own top and bottom
      // margins override the table defaults for that cell. In the projected
      // ordinary row the orphan continuation is the segment's structural owner
      // of its grid area (library projected-segment contract), so an
      // auto-height row (§17.4.80 trHeight) must hold at least those two
      // margins, even though its suppressed content contributes no text. No
      // native paragraph-mark height is assumed for that empty structural
      // cell: the bound is the authored margins alone.
      const TOP = 18;
      const BOTTOM = 14;
      const margined = (carrier: FramePr | null) =>
        ({ ...continuation(carrier), marginTop: TOP, marginBottom: BOTTOM });
      // The host lies far below the ordinary row and AFTER (precondition
      // below), so frame admission cannot move them and only the margin
      // growth is compared. Same carried x and grid as `first`.
      const away = frame({ x: 144, y: 210, w: 200 });
      const row1Control = document([plainRow1()]);
      const [row1Table] = tables(row1Control);
      // Precondition: the margins exceed the sibling-driven control row, so
      // a row that ignores them is visibly too short.
      expect(TOP + BOTTOM).toBeGreaterThan(row1Table!.flowBounds.heightPt);

      const result = document([
        row([restart(away), cell([paragraph('c1b') as CellElement])]),
        row([margined(null), cell([paragraph('c2b') as CellElement])]),
      ]);
      expect(tables(result)).toHaveLength(2);
      expect(ownerOf(result, 0)).toMatchObject({ ordinaryFlow: false, flowBounds: { xPt: 144, yPt: away.y! } });
      const ordinary = ownerOf(result, 1);
      expect(ordinary.ordinaryFlow).toBe(true);
      expect(paths(ordinary)).toEqual([[1, 1]]);
      expect(ordinary.rows).toHaveLength(1);
      expect(ordinary.rows[0]!.heightPt).toBeGreaterThanOrEqual(TOP + BOTTOM);
      expect(ordinary.flowBounds.heightPt).toBeGreaterThanOrEqual(TOP + BOTTOM);
      // Same origin and width as the control; only the height grows.
      const { heightPt: _, ...origin } = row1Table!.flowBounds;
      expect(ordinary.flowBounds).toMatchObject(origin);
      const growth = ordinary.flowBounds.heightPt - row1Table!.flowBounds.heightPt;
      const boxes = runs(result);
      const controlRuns = runs(row1Control);
      // Precondition: AFTER's 10pt line band after the growth stays above
      // the host, so no notBeside clearance enters the comparison.
      expect(at(controlRuns, 'after').y + growth + 10).toBeLessThanOrEqual(away.y!);
      // The top-aligned sibling keeps its place; AFTER follows the taller row.
      for (const text of ['before', 'c2b']) {
        expect(point(at(boxes, text))).toEqual(point(at(controlRuns, text)));
      }
      const after = at(controlRuns, 'after');
      expect(point(at(boxes, 'after'))).toEqual({ text: 'after', x: after.x, y: after.y + growth });
      expectMarkers(boxes);

      // Same-run control: an equal continuation frame keeps the ordinary
      // §17.4.84 merge, and the continuation's margins leave that merged
      // geometry exactly as without them.
      const plain = document(mergedRows(first, first));
      const withMargins = document([
        row([restart(first), cell([paragraph('c1b') as CellElement])]),
        row([margined(first), cell([paragraph('c2b') as CellElement])]),
      ]);
      expect(tables(withMargins)).toHaveLength(1);
      const geometry = (doc: DocumentLayout) => tables(doc).map((node) => ({
        flowBounds: node.flowBounds,
        rows: node.rows.map((laidOut) => laidOut.heightPt),
        borders: node.borders.map((segment) => [segment.from, segment.to, segment.authoredStyle]),
      }));
      expect(geometry(withMargins)).toEqual(geometry(plain));
      expect(runs(withMargins).map(point)).toEqual(runs(plain).map(point));
      expectMarkers(runs(withMargins));
    });
  });

  it('elects by effective positioning and the logical first cell', () => {
    const tableLayout = (wire: object) => ({
      __tableLayout: {
        effectiveStyleId: null,
        grid: { authored: false, columns: [], requiredColumnCount: 0 },
        preferredWidth: null, layout: null, cellSpacing: null,
        ...wire,
      },
    });
    const positioned = {
      leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
      horzAnchor: 'page', horzSpecified: true, vertAnchor: 'page', tblpX: 10, tblpY: 10,
    };
    // Lexical tblpPr whose effective positioning is null still elects.
    const ineffective = layout([table([markerRow('r1', { block: 0, framePr: frame({ x: 150, y: 200 }) })], {
      tblpPr: positioned, ...tableLayout({ ordinaryFlow: true }),
    } as Partial<DocTable>)]);
    expect(tables(ineffective)[0]!.flowBounds).toMatchObject({ xPt: 150, yPt: 200 });

    // bidiVisual: the logical first cell (cells[0]) elects even though it is
    // painted rightmost; a carrier only in logical cell 1 elects nothing.
    const bidi = (carrierCell: 0 | 1) => {
      const authored = markerRow('r1', null);
      authored.cells[carrierCell] = cell([framed('carrier', frame({ x: 150, y: 200 }))]);
      return layout([table([authored], { bidiVisual: true })]);
    };
    expect(tables(bidi(0))[0]!.flowBounds).toMatchObject({ xPt: 150, yPt: 200 });
    const unelected = tables(bidi(1))[0]!;
    expect(unelected.ordinaryFlow).toBe(true);
    expect(unelected.flowBounds.yPt).toBe(MARGIN);
  });

  it('clips a narrow explicit width without shrinking the grid and excludes w+hSpace by max(h,T)', () => {
    // Text-anchored host at the natural cursor; page x=60, w=100 over the
    // natural 200pt grid; exact h=20 is smaller than the host extent T.
    const narrow = frame({
      hAnchor: 'page', x: 60, vAnchor: 'text', y: 0, w: 100,
      hRule: 'exact', h: 20, wrap: 'around', hSpace: 10,
    });
    const result = layout([
      table([markerRow('r1', { block: 0, framePr: narrow })]),
      // Fits only the right gap [170,360] that a 100pt viewport leaves.
      paragraph('f'.repeat(15)) as BodyElement,
      // Fits no gap beside the host, so it starts below the host extent.
      paragraph('g'.repeat(25)) as BodyElement,
    ]);
    const [host] = tables(result);
    const boxes = runs(result);
    expect(host!.flowBounds).toMatchObject({ xPt: 60, yPt: MARGIN, widthPt: 200 });
    expect(host!.columnWidthsPt).toEqual([100, 100]);
    const T = host!.flowBounds.heightPt;
    expect(T).toBeGreaterThan(20);
    // Paint clip: trailing edge at P.x + w; no vertical host clip.
    const clip = host!.clipBounds as LayoutRect;
    expect(clip.xPt + clip.widthPt).toBe(160);
    expect(clip.yPt).toBeLessThanOrEqual(host!.inkBounds.yPt);
    expect(clip.yPt + clip.heightPt).toBeGreaterThanOrEqual(host!.inkBounds.yPt + host!.inkBounds.heightPt);
    // The clipped second cell remains retained, indexed and searchable.
    expect(at(boxes, 'r1b').x).toBeGreaterThanOrEqual(160);
    // External followers: one beside the 100+hSpace exclusion, one below
    // the larger of the exact frame height and the retained host extent.
    expect(at(boxes, 'f'.repeat(15))).toMatchObject({ x: 170, y: MARGIN });
    expect(at(boxes, 'g'.repeat(25))).toMatchObject({ x: MARGIN, y: MARGIN + T });

    // Omitted w: the viewport is the grid and nothing is clipped; an exact
    // row keeps its own §17.4.80 cell clip inside the host.
    const open = layout([table([markerRow('r1', {
      block: 0, framePr: frame({ x: 150, y: 200 }),
    }, { rowHeight: 15, rowHeightRule: 'exact' })])]);
    const [openHost] = tables(open);
    expect(openHost!.flowBounds).toMatchObject({ xPt: 150, yPt: 200 });
    expect(openHost!.clipBounds).toBeUndefined();
    expect(openHost!.rows[0]!.cells[0]!.clipBounds).toBeDefined();
  });

  it('inherits the page clamp: a bottom overflow shifts up, an oversized host pins and splits', () => {
    const shifted = layout([table([markerRow('r1', { block: 0, framePr: frame({ x: 150, y: 380 }) })])]);
    const [host] = tables(shifted);
    expect(host!.flowBounds.yPt + host!.flowBounds.heightPt).toBe(PAGE);

    const tall = Array.from({ length: 16 }, (_, index) =>
      markerRow(`t${index}`, { block: 0, framePr: frame({ x: 150, y: 90 }) }));
    const oversized = layout([table(tall)]);
    const [pinned] = tables(oversized, 0);
    expect(pinned!.flowBounds).toMatchObject({ xPt: 150, yPt: 0 });
    expect(oversized.pages.length).toBeGreaterThan(1);
    // Every source row is emitted exactly once as a source occurrence.
    const sourceRows = oversized.pages.flatMap((_, pageIndex) => tables(oversized, pageIndex)
      .flatMap((fragment) => fragment.rows
        .filter((laidOut) => laidOut.ownership === 'source' && laidOut.fragmentIndex === 0)
        .map((laidOut) => laidOut.logicalRowIndex)));
    expect(sourceRows).toEqual(tall.map((_, index) => index));
  });

  it('discards a host rejected by footnote reserve and places it once on the next page', () => {
    const note: DocNote = { id: '1', content: [paragraph('note') as BodyElement] };
    const referenced = paragraph('n1a', {
      framePr: frame({ hAnchor: 'page', x: 60, vAnchor: 'text', y: 0, wrap: 'around' }),
    });
    referenced.runs.push({ ...referenced.runs[0]!, text: '1', noteRef: { kind: 'footnote', id: '1' } } as never);
    // 30 lines leave a 20pt band: the one-line host fits it alone, but not
    // together with its 16pt note reserve (6pt separator + one 10pt line).
    const filler = Array.from({ length: 30 }, (_, index) => paragraph(`p${index}`) as BodyElement);
    const result = layout([
      ...filler,
      table([row([cell([referenced as CellElement]), cell([paragraph('n1b') as CellElement])], {
        cantSplit: true,
      })]),
      paragraph('after') as BodyElement,
    ], { footnotes: [note] });

    expect(tables(result, 0)).toHaveLength(0);
    // Page 1 text is not wrapped by a rejected host delta.
    expect(runs(result, 0).every((box) => box.x === MARGIN)).toBe(true);
    const moved = tables(result, 1);
    expect(moved).toHaveLength(1);
    expect(moved[0]!.flowBounds).toMatchObject({ xPt: 60, yPt: MARGIN });
    expect(runs(result, 1).filter((box) => box.text === 'n1a')).toHaveLength(1);
  });

  it('repeats source headers inside a continued segment; each source row starts once and paints its text once', () => {
    const header = markerRow('h', null, { isHeader: true });
    const host = markerRow('x', { block: 0, framePr: frame({ x: 150, y: 300 }) });
    const ordinary = Array.from({ length: 12 }, (_, index) => markerRow(`o${index}`, null));
    const result = layout([table([header, host, ...ordinary])]);
    const rowsOn = (pageIndex: number) => tables(result, pageIndex).flatMap((fragment) => fragment.rows);
    // Source row occurrences in page order. A row split across regions keeps
    // its source ownership in every part (fragmentIndex 0, 1, ...); only its
    // first fragment is the row's start.
    const sourceParts = result.pages.flatMap((_, pageIndex) => rowsOn(pageIndex)
      .filter((laidOut) => laidOut.ownership === 'source')
      .map((laidOut) => ({ pageIndex, row: laidOut.logicalRowIndex, fragmentIndex: laidOut.fragmentIndex })));
    const sourceRows = sourceParts.filter((part) => part.fragmentIndex === 0).map((part) => part.row);

    // Source conservation through the painted text index, independent of the
    // row fragment counts below. Non-header rows in source order: the host x,
    // then o0..o11; each holds a1/a2/a3 in cell A and b in cell B.
    const bodyLabels = ['x', ...ordinary.map((_, index) => `o${index}`)];
    const cellParts = ['a1', 'a2', 'a3', 'b'];
    const headerMarkers = cellParts.map((part) => `h${part}`);
    const pageTexts = result.pages.map((_, pageIndex) => runs(result, pageIndex).map((box) => box.text));
    // Page order, then indexed order within a page.
    const sequence = pageTexts.flat().filter((text) => !headerMarkers.includes(text));
    const markerCounts = Object.fromEntries(bodyLabels.flatMap((label) =>
      cellParts.map((part) => [`${label}${part}`, sequence.filter((text) => text === `${label}${part}`).length])));
    // Every non-header marker is painted exactly once across all pages.
    for (const [text, count] of Object.entries(markerCounts)) expect(count, text).toBe(1);
    expect(sequence).toHaveLength(bodyLabels.length * cellParts.length);
    // Rows first appear in source order (a split may place a completed
    // cell B before a later cell-A line, so only first appearances count).
    const firstAppearance = bodyLabels.map((label) =>
      sequence.findIndex((text) => text.startsWith(label) && cellParts.includes(text.slice(label.length))));
    expect(firstAppearance.every((position) => position >= 0)).toBe(true);
    expect(firstAppearance).toEqual([...firstAppearance].sort((a, b) => a - b));
    // Within each original cell A, a1/a2/a3 keep their order.
    for (const label of bodyLabels) {
      const positions = ['a1', 'a2', 'a3'].map((part) => sequence.indexOf(`${label}${part}`));
      expect(positions, label).toEqual([...positions].sort((a, b) => a - b));
    }
    // Repeated headers paint whole, once per region: every page that holds
    // the table shows its header row exactly once (the source row on the
    // first page, a repeated-header occurrence leading each continued page),
    // and holds exactly one complete header copy.
    for (const [pageIndex, texts] of pageTexts.entries()) {
      const pageRows = rowsOn(pageIndex);
      const regions = pageRows.length > 0 ? 1 : 0;
      expect(pageRows.filter((laidOut) => laidOut.logicalRowIndex === 0), `page ${pageIndex} header rows`)
        .toHaveLength(regions);
      for (const marker of headerMarkers) {
        expect(texts.filter((candidate) => candidate === marker).length, `page ${pageIndex} ${marker}`)
          .toBe(regions);
      }
      if (pageIndex === 0) {
        expect(pageRows.find((laidOut) => laidOut.logicalRowIndex === 0)?.ownership).toBe('source');
      } else if (regions > 0) {
        expect(tables(result, pageIndex)[0]!.rows[0])
          .toMatchObject({ logicalRowIndex: 0, ownership: 'repeated-header' });
      }
    }

    // Every source row starts once, in source order.
    expect(sourceRows).toEqual([0, 1, ...ordinary.map((_, index) => index + 2)]);
    // A split row's parts are numbered 0, 1, ... in page order, each on a
    // later page than the part before it; the markers above are painted once,
    // so the parts hold disjoint source text.
    for (const rowIndex of new Set(sourceParts.map((part) => part.row))) {
      const parts = sourceParts.filter((part) => part.row === rowIndex);
      expect(parts.map((part) => part.fragmentIndex), `row ${rowIndex}`).toEqual(parts.map((_, index) => index));
      for (let index = 1; index < parts.length; index += 1) {
        expect(parts[index]!.pageIndex, `row ${rowIndex}`).toBeGreaterThan(parts[index - 1]!.pageIndex);
      }
    }
    expect(tables(result, 1)[0]!.rows[0]).toMatchObject({ logicalRowIndex: 0, ownership: 'repeated-header' });
    // The host keeps its authored page-1 frame position and advances no flow;
    // the following ordinary rows are placed by the ordinary flow's
    // float-exclusion admission against the host's frame exclusion.
    const page1 = tables(result, 0);
    expect(page1.find((fragment) => fragment.rows.some((laidOut) => laidOut.logicalRowIndex === 1))!.flowBounds)
      .toMatchObject({ xPt: 150, yPt: 300 });
  });

  const occurrences = (result: DocumentLayout) => result.pages.flatMap((_, pageIndex) =>
    tables(result, pageIndex).flatMap((fragment) => fragment.rows
      // A row split across regions keeps one first (fragmentIndex 0) source
      // occurrence; its continuation fragments are not new occurrences.
      .filter((laidOut) => laidOut.fragmentIndex === 0)
      .map((laidOut) => ({
        pageIndex, row: laidOut.logicalRowIndex, ownership: laidOut.ownership,
      }))));

  it('repeats every source header row in a host continuation whose run starts inside the headers', () => {
    // Two leading source headers: h0 ordinary, h1 the first row of a
    // page-anchored run whose extent exceeds the page, so the run continues.
    const runFrame = frame({ x: 150, y: 40 });
    const result = layout([table([
      markerRow('h0', null, { isHeader: true }),
      markerRow('h1', { block: 0, framePr: runFrame }, { isHeader: true }),
      ...Array.from({ length: 16 }, (_, index) => markerRow(`x${index}`, { block: 0, framePr: runFrame })),
    ])]);
    const seen = occurrences(result);
    expect(seen.filter((item) => item.ownership === 'source').map((item) => item.row))
      .toEqual(Array.from({ length: 18 }, (_, index) => index));
    const continued = tables(result, 1);
    expect(continued).toHaveLength(1);
    expect(continued[0]!.rows.slice(0, 3).map((laidOut) => [laidOut.logicalRowIndex, laidOut.ownership]))
      .toEqual([[0, 'repeated-header'], [1, 'repeated-header'], [expect.any(Number), 'source']]);
    // Each header occurs once per region: no overlapping duplicate copy.
    expect(continued[0]!.rows.filter((laidOut) => laidOut.logicalRowIndex <= 1)).toHaveLength(2);
  });

  it('enters the next segment as a fresh region when a footnote reserve pushes it off a page', () => {
    // Two columns: column 1 is filled, so a note reserve for the host that
    // follows a completed ordinary segment in column 2 invades committed
    // column-1 content and the paginator restarts the host on the next page.
    const note: DocNote = { id: '1', content: [paragraph('note') as BodyElement] };
    const referenced = paragraph('n1a', {
      framePr: frame({ hAnchor: 'text', x: 0, vAnchor: 'text', y: 0, wrap: 'around' }),
    });
    referenced.runs.push({ ...referenced.runs[0]!, text: '1', noteRef: { kind: 'footnote', id: '1' } } as never);
    const filler = Array.from({ length: 34 }, (_, index) => paragraph(`p${index}`) as BodyElement);
    const result = layout([
      ...filler,
      table([
        markerRow('h', null, { isHeader: true }),
        row([cell([referenced as CellElement]), cell([paragraph('n1b') as CellElement])]),
      ], { colWidths: [60, 60] } as Partial<DocTable>),
    ], {
      footnotes: [note],
      section: section({ columns: { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] } }),
    });
    const seen = occurrences(result);
    expect(seen.filter((item) => item.ownership === 'source').map((item) => [item.pageIndex, item.row]))
      .toEqual([[0, 0], [1, 1]]);
    // The host restarted in a new region and repeats the source header.
    expect(tables(result, 1)[0]!.rows.map((laidOut) => [laidOut.logicalRowIndex, laidOut.ownership]))
      .toEqual([[0, 'repeated-header'], [1, 'source']]);
  });

  it('places a header-story host at its frame with zero story advance and story wrap', () => {
    const headerTable = (carrier: FramePr | null) => table([row([
      cell([carrier ? framed('hc', carrier) : paragraph('hc') as CellElement]),
      cell([paragraph('hd') as CellElement]),
    ])], { colWidths: [50, 50] } as Partial<DocTable>);
    const story = (carrier: FramePr | null) => layout([paragraph('body') as BodyElement], {
      header: [headerTable(carrier), paragraph('hdr') as BodyElement],
    });
    const control = runs(story(null));
    const hosted = runs(story(frame({ hAnchor: 'page', x: 200, vAnchor: 'text', y: 0, wrap: 'around' })));
    const dx = 200 - MARGIN;
    for (const text of ['hc', 'hd']) {
      const before = at(control, text);
      expect(at(hosted, text)).toEqual({ ...before, x: before.x + dx });
    }
    // HDR starts at the host's own story cursor, beside the host exclusion.
    expect(at(hosted, 'hdr')).toMatchObject({ x: MARGIN, y: at(control, 'hc').y - TOP_RULE_INSET });
  });

  it('keeps page and margin anchors of header and footer hosts on the page, not the story band', () => {
    // A page/margin result is the page's. The story is laid out before its
    // band translates it, so the host owns its page axes and the band
    // translation (and the band's own flow extent) leaves it in place.
    const bands = section({ headerDistance: 20, footerDistance: 20 });
    const hostTable = (carrier: FramePr | null) => table([row([
      cell([carrier ? framed('hc', carrier) : paragraph('hc') as CellElement]),
      cell([paragraph('hd') as CellElement]),
    ])], { colWidths: [50, 50] } as Partial<DocTable>);
    const header = (carrier: FramePr | null) => runs(layout([paragraph('body') as BodyElement], {
      section: bands,
      header: [hostTable(carrier), paragraph('hdr') as BodyElement],
    }));
    const control = header(null);
    // The carrier-free table starts the header band at headerDistance.
    const tableTopToText = at(control, 'hc').y - 20;
    for (const [anchor, y, top] of [['page', 300, 0], ['margin', 10, MARGIN]] as const) {
      const hosted = header(frame({ hAnchor: 'page', x: 200, vAnchor: anchor, y, wrap: 'none' }));
      expect(at(hosted, 'hc')).toMatchObject({ x: at(control, 'hc').x + 200 - MARGIN, y: top + y + tableTopToText });
      expect(at(hosted, 'hd').y).toBe(top + y + tableTopToText);
      // Zero story advance; the band still starts at headerDistance.
      expect(at(hosted, 'hdr')).toMatchObject({ x: MARGIN, y: at(control, 'hc').y - TOP_RULE_INSET });
    }

    const footer = (carrier: FramePr | null) => runs(layout([paragraph('body') as BodyElement], {
      section: bands,
      footer: carrier === null
        ? [paragraph('ftr') as BodyElement]
        : [paragraph('ftr') as BodyElement, hostTable(carrier)],
    }));
    const footerControl = footer(null);
    const hostedFooter = footer(frame({ hAnchor: 'page', x: 200, vAnchor: 'page', y: 5, wrap: 'none' }));
    // A footer bottom-aligns its flow extent; the page-owned host adds none.
    expect(at(hostedFooter, 'ftr')).toEqual(at(footerControl, 'ftr'));
    expect(at(hostedFooter, 'hc').y).toBe(5 + tableTopToText);
  });

  it('wraps header text where a page-anchored story host is painted, not one band offset away', () => {
    // Host page/page at y = headerDistance (20), 100pt wide, wrap around: the
    // header paragraph's first line (also at page y 20) must wrap beside it.
    const bands = section({ headerDistance: 20 });
    const result = layout([paragraph('body') as BodyElement], {
      section: bands,
      header: [
        table([row([cell([framed('hc', frame({ x: MARGIN, y: 20, w: 100, wrap: 'around' }))]),
          cell([paragraph('hd') as CellElement])])], { colWidths: [50, 50] } as Partial<DocTable>),
        paragraph('hdr') as BodyElement,
      ],
    });
    const boxes = runs(result);
    expect(at(boxes, 'hc').y).toBeGreaterThanOrEqual(20);
    expect(at(boxes, 'hdr')).toMatchObject({ x: MARGIN + 100, y: 20 });
  });

  it('gives a cut cell paragraph the PAGE field value of the page it is painted on', () => {
    // A long cell paragraph that ends in a PAGE field is cut onto page 2, and
    // its page-2 lines show page 2's value. A bold field result is its own
    // placement; a same-format one joins the preceding text's glyph sequence
    // as one of its source-run owners, which the cell's page-dependence test
    // used to miss (page 1's value carried).
    const pageField = (bold: boolean) => ({
      type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '?',
      bold, italic: false, underline: false, strikethrough: false, fontSize: 10,
      color: null, fontFamily: 'NotInMetrics', background: null, vertAlign: null,
    } as unknown as DocParagraph['runs'][number]);
    const long = (field: DocParagraph['runs'][number] | null) => {
      const authored = paragraph(Array.from({ length: 400 }, () => 'ab').join(' ') + ' ');
      if (field) authored.runs.push(field);
      return authored as CellElement;
    };
    const document = (field: DocParagraph['runs'][number] | null) => layout([table([row([cell([
      long(field),
    ])])], { colWidths: [320] } as Partial<DocTable>)]);
    const own = (boxes: readonly Box[]) => boxes.filter((box) => box.path.join() === '0,0,0,0');
    for (const bold of [true, false]) {
      const label = `bold ${bold}`;
      const result = document(pageField(bold));
      // Precondition: the paragraph is cut onto page 2.
      expect(own(runs(result, 1)).length, label).toBeGreaterThan(0);
      // The field is painted once, on page 2, with page 2's value.
      const fieldText = (boxes: readonly Box[]) => own(boxes)
        .filter((box) => /^\d+$/.test(box.text)).map((box) => box.text);
      expect(fieldText(runs(result, 0)), label).toEqual([]);
      expect(fieldText(runs(result, 1)), label).toEqual(['2']);
      // Page 2 holds the same lines as the field-free document (the field
      // only appends to the last line).
      const lineTops = (boxes: readonly Box[]) => [...new Set(own(boxes).map((box) => box.y))];
      expect(lineTops(runs(result, 1)), label).toEqual(lineTops(runs(document(null), 1)));
    }
    // Four layouts of a 400-word cell paragraph cut across pages (~3.5 s on an
    // idle machine); the default 5 s budget is too tight on loaded CI runners.
  }, 30_000);

  it('places an upright vertical-section host at its physical frame with zero flow advance', () => {
    const vertical = section({ textDirection: 'tbRl' } as Partial<SectionProps>);
    const document = (carrier: FramePr | null) => layout([
      table([markerRow('v', carrier ? { block: 0, framePr: carrier } : null)]),
      paragraph('after') as BodyElement,
    ], { section: vertical });
    const control = document(null);
    const hosted = document(frame({ x: 150, y: 200 }));
    // The upright control table starts at the physical top margin, right of
    // the logical flow start: physical x = pageWidth − margin − grid.
    const dx = 150 - (PAGE - MARGIN - 200);
    const dy = 200 - MARGIN;
    const controlRuns = runs(control);
    const hostRuns = runs(hosted);
    for (const text of ['va1', 'va2', 'va3', 'vb']) {
      const before = at(controlRuns, text);
      expect(at(hostRuns, text)).toEqual({ ...before, x: before.x + dx, y: before.y + dy });
    }
    // Zero advance, for the host as for the upright control table
    // (`word-vertical-section-upright-block-table`): AFTER starts the vertical
    // flow where the table did, not one grid width (200) later. A rotated
    // paragraph's run bounds stay logical, so `y` here is its logical block
    // position and `x` its inline start, which the control table pushes down.
    expect(at(controlRuns, 'after').y).toBe(MARGIN);
    expect(at(controlRuns, 'after').x).toBeGreaterThan(at(hostRuns, 'after').x);
    expect(at(hostRuns, 'after').y).toBe(MARGIN);
    // The notBeside exclusion reaches the vertical text through the inverse
    // upright projection: a host spanning physical x∈[200,400] blocks the
    // vertical lines over logical block [0,200], so AFTER starts at 200.
    const blocking = runs(document(frame({ x: 200, y: 200 })));
    expect(at(blocking, 'after').y).toBe(PAGE - 200);
  });

  describe('owner segments of an upright table on a bottomToTop section page', () => {
    // A producer's bottomToTop section flow (on its nominal btLr token) has
    // a counter-clockwise frame: logical inline runs
    // physically upward from the bottom edge and logical block runs
    // rightward, so logical y = physical x. The Transitional btLr token keeps
    // the clockwise frame. Physical margins top 30 / bottom 70 keep the two
    // column ends apart; left/right stay 40.
    const nativeSection = () => ({
      ...section({ textDirection: 'btLr', marginTop: 30, marginBottom: 70 } as Partial<SectionProps>),
      __sectionPlacement: { sectionId: 'section:0', sectionFlow: 'bottomToTop' },
    }) as unknown as SectionProps;
    const native = (body: BodyElement[]) => layout(body, { section: nativeSection() });
    const hostedTable = (framePr: FramePr, ...rows: DocTableRow[]) =>
      table([markerRow('h', { block: 0, framePr }), ...rows]);
    // Run origins through their painted transforms, in physical page points.
    const painted = (result: DocumentLayout) => textRunGeometryForPage(result, 0).map((run) => {
      const { a, b, c, d, e, f } = run.pointToPage;
      const { xPt, yPt } = run.placement.bounds;
      return { text: run.placement.text, x: a * xPt + c * yPt + e, y: b * xPt + d * yPt + f };
    });
    const paintedAt = (result: DocumentLayout, text: string) => {
      const matches = painted(result).filter((run) => run.text === text);
      expect(matches, text).toHaveLength(1);
      return { x: matches[0]!.x, y: matches[0]!.y };
    };
    // A rotated body paragraph's run bounds stay logical: y is its block position.
    const logicalBlockStart = (result: DocumentLayout, text: string) => {
      const matches = textRunGeometryForPage(result, 0).filter((run) => run.placement.text === text);
      expect(matches, text).toHaveLength(1);
      return matches[0]!.placement.bounds.yPt;
    };

    it('starts a text-anchored host at the physical top of its column', () => {
      // §17.3.1.11 vAnchor text: the band starts at the current column's
      // physical top. In this frame that is the logical inline END, which for
      // a one-column body is the physical top margin (30), so the host is
      // placed as a page-anchored host at y 30 is.
      const textAnchored = native([hostedTable(frame({ x: 150, y: 0, vAnchor: 'text' })), paragraph('after') as BodyElement]);
      const pageAnchored = native([hostedTable(frame({ x: 150, y: 30 })), paragraph('after') as BodyElement]);
      for (const text of ['ha1', 'ha2', 'ha3', 'hb']) {
        expect(paintedAt(textAnchored, text), text).toEqual(paintedAt(pageAnchored, text));
      }
    });

    it('places an ordinary segment after a host where an unsegmented upright table at its flow position is placed', () => {
      // A host advances the vertical flow by 0, so the ordinary segment that
      // follows it starts at the same flow position as a table holding only
      // that row, and takes the same physical origin in the section's own
      // frame (block start at physical x; the block ends at the column's
      // physical bottom).
      const segmented = native([
        hostedTable(frame({ x: 150, y: 200 }), markerRow('w', null)),
        paragraph('after') as BodyElement,
      ]);
      const unsegmented = native([table([markerRow('w', null)]), paragraph('after') as BodyElement]);
      for (const text of ['wa1', 'wa2', 'wa3', 'wb']) {
        expect(paintedAt(segmented, text), text).toEqual(paintedAt(unsegmented, text));
      }
    });

    it('excludes a notBeside host from the vertical lines over its physical span', () => {
      // Logical block = physical x here: a 200pt host over physical x [0, 200]
      // blocks logical block [0, 200], so AFTER starts at 200; over
      // [200, 400] it leaves AFTER at the logical block start, the physical
      // left margin. The clockwise inverse would swap the two outcomes.
      const after = (xPt: number) => logicalBlockStart(
        native([hostedTable(frame({ x: xPt, y: 200 })), paragraph('after') as BodyElement]),
        'after',
      );
      expect(after(0)).toBe(200);
      expect(after(200)).toBe(MARGIN);
    });
  });

  it('treats a Word-saved split into adjacent one-row tables like the two-row source', () => {
    const wire = (offset: number) => ({
      __tableLayout: {
        effectiveStyleId: 'Grid', ordinaryFlow: true,
        logicalSequenceId: 'table-sequence:0', logicalRowOffset: offset, logicalTotalRows: 2,
        // Authored fixed 2000tw (100pt) tracks, as the source-open grid.
        grid: { authored: true, columns: [{ width: '2000' }, { width: '2000' }], requiredColumnCount: 2 },
        preferredWidth: null, layout: { kind: 'fixed' }, cellSpacing: null,
      },
    }) as Partial<DocTable>;
    const first = frame({ x: 150, y: 150 });
    const second = frame({ x: 60, y: 250 });
    const sourceOpen = layout([
      paragraph('before') as BodyElement,
      table([markerRow('r1', { block: 0, framePr: first }), markerRow('r2', { block: 0, framePr: second })]),
      paragraph('after') as BodyElement,
    ]);
    const saved = layout([
      paragraph('before') as BodyElement,
      table([markerRow('r1', { block: 0, framePr: first })], wire(0)),
      table([markerRow('r2', { block: 0, framePr: second })], wire(1)),
      paragraph('after') as BodyElement,
    ]);
    expect(tables(saved).map((host) => host.flowBounds))
      .toEqual(tables(sourceOpen).map((host) => host.flowBounds));
    expect(tables(saved).map((host) => host.rows.map((laidOut) => laidOut.source.path)))
      .toEqual([[[1, 0]], [[2, 0]]]);
    const strip = (boxes: readonly Box[]) => boxes.map(({ text, x, y }) => ({ text, x, y }));
    expect(strip(runs(saved))).toEqual(strip(runs(sourceOpen)));
  });

  it('keeps a page-anchored host out of the page-positioned tblpPr table prescan', () => {
    // A §17.4.57 page-positioned table makes the paginator plan page-start
    // anchors from accepted page-owned tables; a host is not one of them.
    const floating = table([markerRow('f', null)], {
      tblpPr: {
        leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
        horzAnchor: 'page', horzSpecified: true, vertAnchor: 'page', tblpX: 60, tblpY: 60,
      },
    } as Partial<DocTable>);
    const result = layout([
      paragraph('before') as BodyElement,
      floating,
      table([markerRow('r1', { block: 0, framePr: frame({ x: 150, y: 250 }) })]),
      paragraph('after') as BodyElement,
    ]);
    const placed = tables(result);
    expect(placed.map((node) => node.flowBounds.xPt)).toEqual([60, 150]);
    expect(placed[1]!.flowBounds.yPt).toBe(250);
  });

  it('keeps one host per row for many distinct adjacent carriers in source order', () => {
    const count = 240;
    const rows = Array.from({ length: count }, (_, index) => markerRow(`s${index}`, {
      block: 0, framePr: frame({ x: 40 + (index % 7), y: 40 + (index % 5) }),
    }));
    const result = layout([table(rows), paragraph('after') as BodyElement]);
    const hosts = result.pages.flatMap((_, pageIndex) => tables(result, pageIndex));
    expect(hosts.map((host) => host.rows.map((laidOut) => laidOut.logicalRowIndex)))
      .toEqual(rows.map((_, index) => [index]));
    expect(hosts.map((host) => host.flowBounds.xPt).slice(0, 7)).toEqual([40, 41, 42, 43, 44, 45, 46]);
  });
});

describe('page placement of positioned tables below table cells and in stories', () => {
  // A one-cell nested table holding marker row r1, positioned by tblpPr:
  // page (60, 100) or text (cell content x + 20, anchor top + 5).
  const positionedParent = (anchors: 'page' | 'text', pageTopPt = 100) => table([row([cell([
    table([markerRow('r1', null)]) as unknown as CellElement,
  ])])], {
    colWidths: [210],
    tblpPr: {
      leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
      horzAnchor: anchors, horzSpecified: true, vertAnchor: anchors,
      tblpX: anchors === 'page' ? 60 : 20, tblpY: anchors === 'page' ? pageTopPt : 5,
    },
  } as Partial<DocTable>);
  // A positioned holder whose cell holds a positioned child and its anchor
  // MID. Holder frame: page (50, 60) or text (cell content + 7, anchor + 3).
  const positionedHolder = (inner: BodyElement, anchors: 'page' | 'text') => table([row([cell([
    inner as unknown as CellElement, paragraph('mid') as CellElement,
  ])])], {
    colWidths: [250],
    tblpPr: {
      leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
      horzAnchor: anchors, horzSpecified: true, vertAnchor: anchors,
      tblpX: anchors === 'page' ? 50 : 7, tblpY: anchors === 'page' ? 60 : 3,
    },
  } as Partial<DocTable>);
  const markerTexts = ['r1a1', 'r1a2', 'r1a3', 'r1b'];
  // The page-positioned child as a direct child of a body table cell: its
  // final frame is the page's (60, 100).
  const direct = (pageTopPt = 100) => runs(layout([nestedIn(positionedParent('page', pageTopPt), [320])]));

  it('paints a text-anchored positioned nested table at its anchor and wraps its anchor text', () => {
    // §17.4.57 text/text in a body table cell: the frame is the cell content
    // x plus tblpX and the anchor paragraph (POST, the next paragraph) top
    // plus tblpY. Before the final frame owned text/text occurrences, the
    // table reserved its exclusion but was never placed on a page.
    const control = runs(layout([nestedIn(null, [320])]));
    const result = layout([nestedIn(positionedParent('text'), [320])]);
    const frameXPt = at(control, 'pre').x + 20;
    const frameYPt = at(control, 'post').y + 5;
    const [outer] = tables(result);
    expect(outer!.resolvedFloatingTables?.map((placement) => [placement.xPt, placement.yPt]))
      .toEqual([[frameXPt, frameYPt]]);
    // Its rows are painted inside that frame, as the same table without
    // tblpPr in the body is painted inside its own flow position.
    const { tblpPr: _, ...inFlow } = positionedParent('text') as DocTable;
    const bodyControl = layout([inFlow as unknown as BodyElement]);
    const [bodyTable] = tables(bodyControl);
    const boxes = runs(result);
    for (const text of markerTexts) {
      const inBody = at(runs(bodyControl), text);
      expect(at(boxes, text)).toMatchObject({
        x: frameXPt + inBody.x - bodyTable!.flowBounds.xPt,
        y: frameYPt + inBody.y - bodyTable!.flowBounds.yPt,
      });
    }
    // POST, its anchor, wraps beside the 210pt frame (20pt is too narrow).
    expect(at(boxes, 'post')).toMatchObject({ x: frameXPt + 210, y: at(control, 'post').y });
  });

  it('places the positioned children of a positioned table and of an in-flow nested table', () => {
    // Only a paginated table's direct positioned children used to be given a
    // final frame: these were laid out but never placed on a page.
    // Page-anchored grandchild: the same page frame as a direct child.
    const page = direct();
    const pageChild = positionedParent('page');
    for (const holder of [
      positionedHolder(pageChild, 'page'),
      positionedHolder(pageChild, 'text'),
      nestedIn(pageChild, [300]),
    ]) {
      const boxes = runs(layout([nestedIn(holder, [320])]));
      for (const text of markerTexts) {
        expect(at(boxes, text)).toMatchObject({ x: at(page, text).x, y: at(page, text).y });
      }
    }
    // Text-anchored grandchild: its holder's cell content x + 20 and its
    // anchor's top + 5. The child's content sits inside its frame as the
    // direct text/text child's sits inside its own (derived from that case).
    const control = runs(layout([nestedIn(null, [320])]));
    const directText = runs(layout([nestedIn(positionedParent('text'), [320])]));
    const inside = (text: string) => ({
      x: at(directText, text).x - (at(control, 'pre').x + 20),
      y: at(directText, text).y - (at(control, 'post').y + 5),
    });
    const textChild = positionedParent('text');
    // Page holder at (50, 60): cell content x 50 + 5; MID, the anchor, starts
    // the cell at y 60, below the holder's top rule.
    const inPageHolder = runs(layout([nestedIn(positionedHolder(textChild, 'page'), [320])]));
    for (const text of markerTexts) {
      expect(at(inPageHolder, text)).toMatchObject({
        x: 50 + 5 + 20 + inside(text).x, y: 60 + TOP_RULE_INSET + 5 + inside(text).y,
      });
    }
    // In-flow holder: cell content x is the inner PRE's x; POST (the anchor)
    // starts below PRE. The anchor wraps beside the 210pt child.
    const inFlow = runs(layout([nestedIn(nestedIn(textChild, [300]), [320])]));
    const innerPre = inFlow.filter((box) => box.text === 'pre')[1]!;
    for (const text of markerTexts) {
      expect(at(inFlow, text)).toMatchObject({
        x: innerPre.x + 20 + inside(text).x, y: innerPre.y + 10 + 5 + inside(text).y,
      });
    }
    expect(inFlow.filter((box) => box.text === 'post')[0]).toMatchObject({
      x: innerPre.x + 20 + 210, y: innerPre.y + 10,
    });
  });

  it('places a page-positioned table through positioned tables nested in one another', () => {
    // Each positioned level resolves its final frame by an exact solve whose
    // passes lay out the levels inside it (table-pagination.ts
    // resolveFinalFrameChild, each nested layout charged to the session
    // budget). Text-anchored holders keep each level inside the page; the
    // innermost page-positioned table keeps the page's frame at every depth.
    let element = positionedParent('page');
    for (let level = 0; level < 3; level += 1) element = positionedHolder(element, 'text');
    const boxes = runs(layout([nestedIn(element, [320])]));
    const page = direct();
    for (const text of markerTexts) {
      expect(at(boxes, text), text).toMatchObject({ x: at(page, text).x, y: at(page, text).y });
    }
  });

  it('places positioned tables of header and footnote tables on the page', () => {
    // Story tables used to lay out their §17.4.57 positioned children
    // without ever placing them: their rows were not painted at all. A
    // footnote is placed through the band plan its composed position states.
    const inBody = direct();
    const control = runs(layout([nestedIn(null, [320])]));
    const directText = runs(layout([nestedIn(positionedParent('text'), [320])]));
    for (const where of ['header', 'footnote'] as const) {
      const story = (anchors: 'page' | 'text') => {
        const content = [nestedIn(positionedParent(anchors), [320])];
        const body = paragraph('body');
        if (where === 'footnote') {
          body.runs.push({ ...body.runs[0]!, text: '1', noteRef: { kind: 'footnote', id: 'f1' } } as never);
        }
        return layout([body as BodyElement], {
          section: section({ headerDistance: 20 }),
          ...(where === 'header' ? { header: content } : { footnotes: [{ id: 'f1', content }] }),
        });
      };
      // Page-positioned: the same page frame as in a body table.
      const page = runs(story('page'));
      for (const text of markerTexts) {
        expect(at(page, text), `${where} page ${text}`).toMatchObject({ x: at(inBody, text).x, y: at(inBody, text).y });
      }
      // Text-positioned: tblpX 20 / tblpY 5 from the story cell and its
      // anchor POST, the content inside its frame as in the body.
      const text = runs(story('text'));
      for (const marker of markerTexts) {
        expect(at(text, marker), `${where} text ${marker}`).toMatchObject({
          x: at(text, 'pre').x + 20 + at(directText, marker).x - (at(control, 'pre').x + 20),
          y: at(text, 'post').y + 5 + at(directText, marker).y - (at(control, 'post').y + 5),
        });
      }
    }
  });

  it('places positioned tables of a footnote table on the page when footnotes may continue', () => {
    // Footnote continuation is the public default. A note whose table holds
    // a positioned child is then a destination-dependent full acquisition:
    // it is never retained per page, and each band it is laid out for is a
    // charged trial. It must still be laid out with the band its composed
    // position states, exactly as when continuation is disabled (the case
    // above): the page anchor at the body's page frame, the text anchor at
    // tblpX 20 / tblpY 5 from its story cell and its anchor POST.
    const inBody = direct();
    const control = runs(layout([nestedIn(null, [320])]));
    const directText = runs(layout([nestedIn(positionedParent('text'), [320])]));
    const footnote = (anchors: 'page' | 'text') => {
      const body = paragraph('body');
      body.runs.push({ ...body.runs[0]!, text: '1', noteRef: { kind: 'footnote', id: 'f1' } } as never);
      return runs(layout([body as BodyElement], {
        section: section({ headerDistance: 20 }),
        footnotes: [{ id: 'f1', content: [nestedIn(positionedParent(anchors), [320])] }],
        allowFootnoteContinuation: true,
      }));
    };
    const page = footnote('page');
    for (const text of markerTexts) {
      expect(at(page, text), `page ${text}`).toMatchObject({ x: at(inBody, text).x, y: at(inBody, text).y });
    }
    const text = footnote('text');
    for (const marker of markerTexts) {
      expect(at(text, marker), `text ${marker}`).toMatchObject({
        x: at(text, 'pre').x + 20 + at(directText, marker).x - (at(control, 'pre').x + 20),
        y: at(text, 'post').y + 5 + at(directText, marker).y - (at(control, 'post').y + 5),
      });
    }
  });

  it('places positioned tables of vertical-section footnote and endnote tables in the note page frame', () => {
    // Ordinary §17.4.57 placement only (row-owner framePr in notes is inert).
    // A tbRl section lays its notes out in logical coordinates and paints
    // them through its logical-to-physical transform (physical x = page width
    // − logical y, physical y = logical x). Library policy, not Word
    // evidence: a note table's positioned child resolves against the physical
    // page and margin boxes carried into those coordinates, so it moves with
    // its operands along the note's axes — tblpX (logical inline) physically
    // down, tblpY (logical block) physically left — and from the page band to
    // the margin band by the margin box's logical origin (top margin, right
    // margin), physically (−right, +top). A non-square page and four distinct
    // margins keep each axis and edge apart.
    const vertical = section({
      textDirection: 'tbRl', pageWidth: 400, pageHeight: 500,
      marginTop: 40, marginRight: 60, marginBottom: 50, marginLeft: 30,
    } as Partial<SectionProps>);
    const child = (anchors: 'page' | 'margin', xPt: number, yPt: number) => table([
      row([cell([paragraph('vc') as CellElement])]),
    ], {
      colWidths: [50], borders: noBorders, cellMarginLeft: 0, cellMarginRight: 0,
      tblpPr: {
        leftFromText: 0, rightFromText: 0, topFromText: 0, bottomFromText: 0,
        horzAnchor: anchors, horzSpecified: true, vertAnchor: anchors, tblpX: xPt, tblpY: yPt,
      },
    } as Partial<DocTable>);
    // Each run's bounds origin through its painted transform, on any page.
    const painted = (result: DocumentLayout, text: string) => {
      const matches = result.pages.flatMap((_, pageIndex) => textRunGeometryForPage(result, pageIndex))
        .filter((run) => run.placement.text === text);
      expect(matches, text).toHaveLength(1);
      const { pointToPage: { a, b, c, d, e, f }, placement: { bounds } } = matches[0]!;
      return { x: a * bounds.xPt + c * bounds.yPt + e, y: b * bounds.xPt + d * bounds.yPt + f };
    };
    for (const kind of ['footnote', 'endnote'] as const) {
      // The holder table's cell holds the positioned child and its anchor NA;
      // `moved` adds a note paragraph before and after the holder, which
      // moves it within the note and the note on the page.
      const document = (positioned: BodyElement, moved: boolean) => {
        const body = paragraph('body');
        body.runs.push({ ...body.runs[0]!, text: '1', noteRef: { kind, id: 'n1' } } as never);
        const holder = table([row([cell([positioned as unknown as CellElement, paragraph('na') as CellElement])])], {
          colWidths: [100], borders: noBorders,
        } as Partial<DocTable>);
        const content = moved
          ? [paragraph('lead') as BodyElement, holder, paragraph('tail') as BodyElement]
          : [holder];
        const notes = [{ id: 'n1', content }];
        return layout([body as BodyElement], {
          section: vertical,
          ...(kind === 'footnote' ? { footnotes: notes } : { endnotes: notes }),
        });
      };
      const base = document(child('page', 100, 150), false);
      const page = painted(base, 'vc');
      // Page-anchored: where the holder sits in the note does not move it.
      const moved = document(child('page', 100, 150), true);
      expect(painted(moved, 'na'), `${kind} holder moved`).not.toEqual(painted(base, 'na'));
      expect(painted(moved, 'vc'), `${kind} page anchor`).toEqual(page);
      // +20 tblpX, +10 tblpY along the note's axes.
      expect(painted(document(child('page', 120, 160), false), 'vc'), `${kind} operands`)
        .toEqual({ x: page.x - 10, y: page.y + 20 });
      // The margin band's logical origin is (top 40, right 60).
      expect(painted(document(child('margin', 100, 150), false), 'vc'), `${kind} margin`)
        .toEqual({ x: page.x - 60, y: page.y + 40 });
    }
  });

  describe('text box stories', () => {
    // A DrawingML-anchored text box: page offset (50, 100), or 10pt below
    // its anchor paragraph; 300×150, wrap none, no insets. Its story holds a
    // table whose cell holds the page-positioned child.
    const missingEdges = {
      topPt: null, topStatus: 'missing', rightPt: null, rightStatus: 'missing',
      bottomPt: null, bottomStatus: 'missing', leftPt: null, leftStatus: 'missing',
    };
    const anchorInput = (heightPt: number, vertical: Record<string, unknown>) => ({
      occurrenceId: 'box',
      simplePosition: { enabled: false, status: 'valid', xPt: 0, xStatus: 'valid', yPt: 0, yStatus: 'valid' },
      horizontal: { relativeFrom: 'page', relativeFromStatus: 'valid', choice: { kind: 'offset', valuePt: 50 } },
      vertical: { relativeFrom: 'page', relativeFromStatus: 'valid', choice: { kind: 'offset', valuePt: 100 }, ...vertical },
      extent: { widthPt: 300, widthStatus: 'valid', heightPt, heightStatus: 'valid' },
      parentEffectExtent: missingEdges,
      anchorDistances: missingEdges,
      relativeSize: { horizontal: null, vertical: null },
      wrap: { kind: 'none', authoredKinds: ['wrapNone'], side: 'bothSides', distances: missingEdges, effectExtent: null, polygon: null },
      behavior: {
        behindDoc: false, behindDocStatus: 'valid', relativeHeight: 1, relativeHeightStatus: 'valid',
        locked: false, lockedStatus: 'valid', allowOverlap: true, allowOverlapStatus: 'valid',
        layoutInCell: false, layoutInCellStatus: 'valid',
      },
      group: null,
    });
    const anchoredParagraph = (
      content: BodyElement[],
      extra: Record<string, unknown> = {},
      vertical: Record<string, unknown> = {},
      occurrenceId = 'box',
    ) => {
      const host = paragraph('p');
      host.runs.push({ type: 'anchorHost', fontSize: 10, anchorOccurrenceId: occurrenceId } as never);
      host.runs.push({
        type: 'shape', widthPt: 300, heightPt: 150, anchorXPt: 50, anchorYPt: 100,
        anchorXFromMargin: false, anchorYFromPara: false, zOrder: 0, subpaths: [],
        presetGeometry: 'rect', fill: null, stroke: null, wrapMode: 'none',
        textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0, textAutofit: 'none',
        textBlocks: [{ text: 'x', fontSizePt: 10 }], textBoxContent: content, ...extra,
        anchorAcquisitionInput: { ...anchorInput(150, vertical), occurrenceId },
      } as never);
      return host;
    };
    const laidOut = (model: Partial<DocxDocumentModel>, allowFootnoteContinuation = false) => {
      const normalized = normalizeInternalDocumentModel({
        section: section({ headerDistance: 20 }),
        body: [],
        headers: { default: null, first: null, even: null },
        footers: { default: null, first: null, even: null },
        footnotes: [], endnotes: [], fontFamilyClasses: {},
        ...model,
      } as unknown as DocxDocumentModel);
      return layoutDocument(
        normalized.document,
        createLayoutServices(normalized.document, {
          measureContext: measureContext(),
          ...(allowFootnoteContinuation ? { allowFootnoteContinuation: true } : {}),
        }),
        { currentDateMs: 0 },
      );
    };
    // Each run's bounds origin through its painted transform.
    const painted = (result: DocumentLayout): readonly Box[] => textRunGeometryForPage(result, 0).map((run) => {
      const { a, b, c, d, e, f } = run.pointToPage;
      const { xPt, yPt } = run.placement.bounds;
      return { text: run.placement.text, x: a * xPt + c * yPt + e, y: b * xPt + d * yPt + f, path: run.source.path };
    });
    const storyContent = () => [nestedIn(positionedParent('page'), [280]), paragraph('tp') as BodyElement];
    const expectAtPageFrame = (result: DocumentLayout, label: string) => {
      const page = direct();
      for (const text of markerTexts) {
        expect(at(painted(result), text), `${label} ${text}`)
          .toMatchObject({ x: at(page, text).x, y: at(page, text).y });
      }
    };

    it('places a positioned table of a text box story table on the page through the box placement', () => {
      // Page-positioned (60, 100) child of a story table cell, in a box the
      // spAutoFit bottom alignment moves after its story is laid out.
      const result = laidOut({
        body: [anchoredParagraph(storyContent(), { textAutofit: 'sp' }, {
          choice: { kind: 'align', value: 'bottom' },
        }) as BodyElement],
      });
      // Precondition: the fitted box moved (its story with it).
      expect(at(painted(result), 'pre').y).not.toBe(100);
      expectAtPageFrame(result, 'autofit');
    });

    it('places it through a header band and through a body table cell', () => {
      // Paragraph-relative boxes ride their paragraph: in a header through
      // the band, in a table cell through the cell's page position, which
      // only the table's pagination knows (table-pagination.ts
      // placeRowNestedContent re-acquires the cell paragraph with it).
      const box = () => anchoredParagraph(storyContent(), {}, {
        relativeFrom: 'paragraph', choice: { kind: 'offset', valuePt: 10 },
      });
      expectAtPageFrame(laidOut({
        body: [paragraph('body') as BodyElement],
        headers: { default: { body: [box()] }, first: null, even: null },
      } as unknown as Partial<DocxDocumentModel>), 'header');
      expectAtPageFrame(laidOut({
        body: [table([row([cell([box() as CellElement])])], { colWidths: [320] } as Partial<DocTable>)],
      }), 'cell');
    });

    it('paints the story of a text box in a body table cell where the text index places it', async () => {
      // The same body-cell case through the production page painter. P, the
      // box's anchor paragraph, is placed by its cell at the cell content
      // origin (40 + 5, 40); the box's story is its own coordinate root, so
      // the cell's placement of P is not undone again inside it. The child's
      // frame is the authored page (60, 100): r1a1 sits at (60 + 5 + 5, 100)
      // inside the child's and the marker table's 5pt left cell margins.
      // Expected paint: each glyph run drawn once, at the page point the
      // retained text index states for its origin, through the transforms
      // the painter actually applies; painting measures no text.
      const result = laidOut({
        body: [table([row([cell([anchoredParagraph(storyContent(), {}, {
          relativeFrom: 'paragraph', choice: { kind: 'offset', valuePt: 10 },
        }) as CellElement])])], { colWidths: [320] } as Partial<DocTable>)],
      });
      const indexed = textRunGeometryForPage(result, 0);
      const pageOrigin = (text: string) => {
        const matches = indexed.filter((run) => run.placement.text === text);
        expect(matches, text).toHaveLength(1);
        const { pointToPage: { a, b, c, d, e, f }, placement: { origin, bounds } } = matches[0]!;
        return {
          origin: { x: a * origin.xPt + c * origin.yPt + e, y: b * origin.xPt + d * origin.yPt + f },
          topLeft: { x: a * bounds.xPt + c * bounds.yPt + e, y: b * bounds.xPt + d * bounds.yPt + f },
          baselinePt: origin.yPt - bounds.yPt,
        };
      };
      // Precondition: P is translated by its cell, so a painter that undid
      // that translation inside the story would move the child.
      // Below the cell's 1pt top rule.
      expect(pageOrigin('p').topLeft).toEqual({ x: 45, y: 40 + TOP_RULE_INSET });
      // Inside the child's and the marker table's top rules.
      expect(pageOrigin('r1a1').topLeft).toEqual({ x: 70, y: 100 + 2 * TOP_RULE_INSET });

      const recorder = recordingCanvas();
      await paintLayoutPage(result, 0, recorder.canvas, { dpr: 1, scale: 1 }, { paint() {} });
      expect(recorder.measures()).toBe(0);
      const paintedAt = (text: string) => {
        const calls = recorder.texts.filter((call) => call.text === text);
        expect(calls, text).toHaveLength(1);
        return { x: calls[0]!.x, y: calls[0]!.y };
      };
      for (const text of ['p', ...markerTexts]) {
        expect(paintedAt(text), text).toEqual(pageOrigin(text).origin);
      }
      expect(paintedAt('r1a1')).toEqual({ x: 70, y: 100 + 2 * TOP_RULE_INSET + pageOrigin('r1a1').baselinePt });
    });

    it('acquires a text box in the continued part of a footnote on that part\'s page', () => {
      // A footnote of 40 one-line paragraphs and a final paragraph anchoring
      // a text box whose story holds PG and a PAGE field. It cannot fit
      // beside its reference, so it continues; its final paragraph is painted
      // on a later page. That part is acquired for its own destination page,
      // and the nested text box story is bound to that note context, so the
      // PAGE field is that page's number. Reusing the reference page's
      // acquisition (or binding the box to it) would paint the first page's.
      const body = paragraph('body');
      body.runs.push({ ...body.runs[0]!, text: '1', noteRef: { kind: 'footnote', id: 'f1' } } as never);
      const fieldParagraph = paragraph('pg');
      fieldParagraph.runs.push({
        ...fieldParagraph.runs[0]!, type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '?',
      } as never);
      const result = laidOut({
        body: [body as BodyElement, ...Array.from({ length: 40 }, () => paragraph('x') as BodyElement)],
        footnotes: [{
          id: 'f1',
          content: [
            ...Array.from({ length: 40 }, () => paragraph('n') as BodyElement),
            anchoredParagraph([fieldParagraph as BodyElement], {}, {
              relativeFrom: 'paragraph', choice: { kind: 'offset', valuePt: 0 },
            }) as BodyElement,
          ],
        }],
      }, true);
      const pages = result.pages.map((_, pageIndex) => textRunGeometryForPage(result, pageIndex));
      const boxPages = pages.flatMap((page, pageIndex) =>
        page.some((run) => run.placement.text === 'pg') ? [pageIndex] : []);
      expect(boxPages).toHaveLength(1);
      const [pageIndex] = boxPages as [number];
      // Precondition: the text box is painted with the continued part.
      expect(pageIndex).toBeGreaterThan(0);
      const owner = pages[pageIndex]!.find((run) => run.placement.text === 'pg')!.source;
      expect(pages[pageIndex]!
        .filter((run) => JSON.stringify(run.source) === JSON.stringify(owner))
        .map((run) => run.placement.text).join(''))
        .toBe(`pg${pageIndex + 1}`);
    });

    it('keeps page-positioned tables fixed across two column bands of one continued note', () => {
      // §17.4.57 page anchoring survives both the band's final top and the
      // continuation's source cut. Each column needs its own plan for f1.
      const body = paragraph('body');
      body.runs.push({ ...body.runs[0]!, text: '1', noteRef: { kind: 'footnote', id: 'f1' } } as never);
      const box = (id: string) => anchoredParagraph([
        nestedIn(positionedParent('page', 110), [280]), paragraph('tp') as BodyElement,
      ], {}, { relativeFrom: 'paragraph', choice: { kind: 'offset', valuePt: -280 } }, id) as BodyElement;
      const result = laidOut({
        section: section({ pageWidth: 760, columns: {
          count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [],
        } }),
        settings: { compatibilityMode: 11 },
        body: [body as BodyElement],
        footnotes: [{ id: 'f1', content: [
          box('head'), ...Array.from({ length: 40 }, () => paragraph('n') as BodyElement), box('tail'),
        ] }],
      }, true);
      const positions = result.pages.flatMap((_, pageIndex) => textRunGeometryForPage(result, pageIndex)
        .filter((run) => run.placement.text === 'r1a1')
        .map((run) => ({ pageIndex, y: run.pointToPage.b * run.placement.bounds.xPt
          + run.pointToPage.d * run.placement.bounds.yPt + run.pointToPage.f })));
      // The table's top border contributes 2pt above its text.
      expect(positions).toEqual([{ pageIndex: 0, y: 112 }, { pageIndex: 0, y: 112 }]);
    });

    it('keeps a page-positioned table of a box story in a continued footnote tail at its page frame', async () => {
      // The same text box chain (story table -> cell -> §17.4.57 tblpPr table)
      // anchored by the last paragraph P of a footnote that cannot fit beside
      // its reference, so P is painted in the continued tail on a later page.
      // The box is page-offset horizontally (x 50) and follows P vertically
      // (RISE above P's top); the tail is cut from a whole-note acquisition,
      // so P reaches the page through its band and its source cut. §17.4.57
      // page/page: the child's frame is the page's (60, 100), as in a body
      // table cell (direct()): r1a1 at (60 + 5 + 5, 100) inside the child's
      // and the marker table's 5pt left cell margins. That point does not
      // depend on the cut or on where the note band lands. A text/text child
      // (cell content x + 20, its anchor POST's top + 5) and the story's own
      // PRE/POST/TP follow the box: each keeps the page x and the offset from
      // P it has when the same box paragraph is laid out in the body. So a
      // tail that cancelled every translation, or moved the page child with
      // the box, fails one of the two. RISE keeps the box over y 100 in both
      // the body control (P at 40 + 10 × CONTROL_LINES) and the tail (P on the
      // band's last line); the expectations do not use either position.
      const RISE = -280;
      const CONTROL_LINES = 26;
      const box = (content: BodyElement[]) => anchoredParagraph(content, {}, {
        relativeFrom: 'paragraph', choice: { kind: 'offset', valuePt: RISE },
      }) as BodyElement;
      const noteTail = (content: () => BodyElement[]) => {
        const body = paragraph('body');
        body.runs.push({ ...body.runs[0]!, text: '1', noteRef: { kind: 'footnote', id: 'f1' } } as never);
        return laidOut({
          body: [body as BodyElement, ...Array.from({ length: 40 }, () => paragraph('x') as BodyElement)],
          footnotes: [{
            id: 'f1',
            content: [...Array.from({ length: 40 }, () => paragraph('n') as BodyElement), box(content())],
          }],
        }, true);
      };
      const bodyControl = (content: () => BodyElement[]) => laidOut({
        body: [...Array.from({ length: CONTROL_LINES }, () => paragraph('x') as BodyElement), box(content())],
      });
      // Every run on every page: bounds and origin through its full transform.
      const indexed = (result: DocumentLayout) => result.pages.flatMap((_, pageIndex) =>
        textRunGeometryForPage(result, pageIndex).map((run) => {
          const { pointToPage: { a, b, c, d, e, f }, placement: { text, origin, bounds } } = run;
          return {
            pageIndex, text,
            x: a * bounds.xPt + c * bounds.yPt + e, y: b * bounds.xPt + d * bounds.yPt + f,
            origin: { x: a * origin.xPt + c * origin.yPt + e, y: b * origin.xPt + d * origin.yPt + f },
          };
        }));
      type Indexed = ReturnType<typeof indexed>[number];
      const once = (boxes: readonly Indexed[], text: string) => {
        const matches = boxes.filter((candidate) => candidate.text === text);
        expect(matches, text).toHaveLength(1);
        return matches[0]!;
      };
      const textStory = () => [nestedIn(positionedParent('text'), [280]), paragraph('tp') as BodyElement];
      const storyTexts = ['pre', 'post', 'tp'];
      // The page child sits clear of the tail's host-following lines, so only
      // the host's movement (not a boundary wrap) is under test.
      const page = direct(110);
      const tailStory = () => [nestedIn(positionedParent('page', 110), [280]), paragraph('tp') as BodyElement];

      for (const [label, content] of [['page', tailStory], ['text', textStory]] as const) {
        const result = noteTail(content);
        const boxes = indexed(result);
        const controlBoxes = indexed(bodyControl(content));
        // Preconditions: the note is really continued (its head is on the
        // reference page, every note line painted once) and P is in the tail.
        const anchor = once(boxes, 'p');
        expect(anchor.pageIndex, label).toBeGreaterThan(0);
        expect(boxes.filter((run) => run.pageIndex === 0 && run.text === 'n').length, label).toBeGreaterThan(0);
        expect(boxes.filter((run) => run.text === 'n'), label).toHaveLength(40);
        // Story markers exactly once, all on the tail page.
        for (const text of [...storyTexts, ...markerTexts]) {
          expect(once(boxes, text).pageIndex, `${label} ${text}`).toBe(anchor.pageIndex);
        }
        // The host really moved against the body control, so host-following
        // content left at its control point fails.
        const controlAnchor = once(controlBoxes, 'p');
        expect(controlAnchor.pageIndex).toBe(0);
        const hostDy = anchor.y - controlAnchor.y;
        expect(hostDy, label).not.toBe(0);
        const following = label === 'page' ? storyTexts : [...storyTexts, ...markerTexts];
        for (const text of following) {
          const control = once(controlBoxes, text);
          expect({ x: once(boxes, text).x, y: once(boxes, text).y }, `${label} ${text}`)
            .toEqual({ x: control.x, y: control.y + hostDy });
        }
        if (label === 'page') {
          // The page child stays at the page frame of the direct body case.
          for (const text of markerTexts) {
            expect({ x: once(boxes, text).x, y: once(boxes, text).y }, `page ${text}`)
              .toEqual({ x: at(page, text).x, y: at(page, text).y });
          }
          expect({ x: once(boxes, 'r1a1').x, y: once(boxes, 'r1a1').y })
            .toEqual({ x: 70, y: 110 + 2 * TOP_RULE_INSET });
        }

        // Production paint of the tail page through its full transforms: each
        // run drawn once, at its indexed origin; painting measures no text.
        const recorder = recordingCanvas();
        await paintLayoutPage(result, anchor.pageIndex, recorder.canvas,
          { dpr: 1, scale: 1 }, { paint() {} });
        expect(recorder.measures(), label).toBe(0);
        for (const text of ['p', ...storyTexts, ...markerTexts]) {
          const calls = recorder.texts.filter((call) => call.text === text);
          expect(calls, `${label} paint ${text}`).toHaveLength(1);
          expect({ x: calls[0]!.x, y: calls[0]!.y }, `${label} paint ${text}`).toEqual(once(boxes, text).origin);
        }
        if (label === 'page') {
          const r1a1 = once(boxes, 'r1a1');
          const call = recorder.texts.find((candidate) => candidate.text === 'r1a1')!;
          expect({ x: call.x, y: call.y }).toEqual({ x: 70, y: 110 + 2 * TOP_RULE_INSET + r1a1.origin.y - r1a1.y });
        }
      }
    });
  });
});

describe('w:framePr wrap (§17.18.104) at the production layout boundary', () => {
  // none and notBeside forbid text beside the frame and resume it on the
  // next line not intersecting the frame's extents; around lets the line use
  // the remaining space. A page-anchored frame at page (150, 40), 200pt wide,
  // over the first body line (x 40–360, y 40–50): 'after' (50pt) fits in the
  // 110pt beside it.
  const at40 = (wrap: FramePr['wrap']) => frame({ x: 150, y: 40, w: 200, wrap });
  const after = (result: DocumentLayout) => at(runs(result), 'after');

  it('puts body text after a none paragraph frame on the next line clear of it', () => {
    const document = (wrap: FramePr['wrap']) => layout([
      paragraph('fr', { framePr: at40(wrap) }) as BodyElement,
      paragraph('after') as BodyElement,
    ]);
    const frameBottomPt = at(runs(document('around')), 'fr').y + 10;
    expect(after(document('around'))).toMatchObject({ x: MARGIN, y: 40 });
    for (const wrap of ['none', 'notBeside'] as const) {
      expect(after(document(wrap)), wrap).toMatchObject({ x: MARGIN, y: frameBottomPt });
    }
  });

  it('puts body text after a none promoted row host on the next line clear of it', () => {
    const document = (wrap: FramePr['wrap']) => layout([
      table([markerRow('r1', { block: 0, framePr: at40(wrap) })]),
      paragraph('after') as BodyElement,
    ]);
    // The host table, borders included, as the around control places it.
    const [host] = tables(document('around'));
    const hostBottomPt = host!.flowBounds.yPt + host!.flowBounds.heightPt;
    expect(after(document('around'))).toMatchObject({ x: MARGIN, y: 40 });
    for (const wrap of ['none', 'notBeside'] as const) {
      expect(after(document(wrap)), wrap).toMatchObject({ x: MARGIN, y: hostBottomPt });
    }
  });
});
