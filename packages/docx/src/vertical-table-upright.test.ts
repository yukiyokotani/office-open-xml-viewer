import { describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import type {
  DocxDocumentModel,
  DocTable,
  DocTableRow,
  DocTableCell,
  DocParagraph,
  SectionProps,
  BodyElement,
} from './types';
import type { TableFragmentLayout } from './layout/table-pagination.js';
import type { DocumentLayout, LayoutPage, PaintNode } from './layout/types.js';
import { testFontSnapshot } from './layout/test-font-snapshot.js';

// ECMA-376 §17.6.20 + §17.4.80/§17.18.37 — issue #988 batch-3 adjudication ④:
// a table CELL inside a vertical (tbRl) section renders like a normal
// HORIZONTAL cell — the section's vertical text direction does NOT propagate
// into the cell:
//   - cell text is laid out horizontally (left→right, wrapping downward),
//   - a fixed `tcW` is the cell's PHYSICAL horizontal width,
//   - `trHeight hRule="exact"` clips overflow at the physical row height,
//   - auto row height GROWS to enclose the content.
// The table block sits upright at the flow position: its physical top edge is
// the top content margin (the column axis start) and it advances the vertical
// flow by its PHYSICAL WIDTH (columns progress right→left past it).
//
// Word ground truth = the batch-3 vertical-table fixture PDF (two 1-cell
// 1-in-wide fixed tables, exact 2 in vs auto): both cells lay their long CJK
// string out horizontally at the physical width; the exact row's border box is
// exactly 2 in tall with lines 10+ clipped; the auto row grew to fit all lines.

/** Recording 2D context (same skeleton as table-clip-exact.test.ts): captures
 *  `rect()` calls so the §17.4.80 exact-row clip band can be asserted, with
 *  deterministic char-count metrics. */
interface RectCall { x: number; y: number; w: number; h: number; }
type PaintEvent =
  | Readonly<{ kind: 'text'; text: string }>
  | Readonly<{ kind: 'stroke'; color: string }>;
function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  rectCalls: RectCall[];
  paintEvents: PaintEvent[];
  measureCalls: () => number;
} {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const rectCalls: RectCall[] = [];
  const paintEvents: PaintEvent[] = [];
  let measures = 0;
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      measures += 1;
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {
      paintEvents.push({ kind: 'stroke', color: String(ctx.strokeStyle) });
    }, fill() {}, fillRect() {},
    strokeRect() {},
    rect(x: number, y: number, w: number, h: number) {
      rectCalls.push({ x, y, w, h });
    },
    clip() {},
    scale() {}, translate() {}, rotate() {}, setTransform() {},
    setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText(text: string) {
      paintEvents.push({ kind: 'text', text });
    }, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0,
    height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return {
    canvas: canvas as unknown as HTMLCanvasElement,
    rectCalls,
    paintEvents,
    measureCalls: () => measures,
  };
}

function emptyBorders() {
  return { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };
}

function bodyParagraph(text: string): DocParagraph {
  return {
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: text === '' ? [] : [{
      type: 'text', text,
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman',
    widowControl: false,
  } as unknown as DocParagraph;
}

/** 1×1 fixed-width (50 pt) table. `exactPt` sets `trHeight hRule="exact"`;
 *  null keeps the row auto. The cell text ('tok tok …') wraps one token per
 *  line at the 50 pt cell width (each 'tok ' measures 40 pt with the fake
 *  10 px-per-char metrics), giving 12 lines ≈ 130+ pt of content — far taller
 *  than an 80 pt exact row. */
function fixedTable(text: string, exactPt: number | null): DocTable {
  const cell: DocTableCell = {
    content: [{ type: 'paragraph', ...bodyParagraph(text) }],
    colSpan: 1,
    vMerge: null,
    borders: emptyBorders(),
    background: null,
    vAlign: 'top',
    widthPt: 50,
    marginTop: 0, marginBottom: 0, marginLeft: 0, marginRight: 0,
  } as unknown as DocTableCell;
  const row: DocTableRow = {
    cells: [cell],
    rowHeight: exactPt,
    rowHeightRule: exactPt == null ? 'auto' : 'exact',
    isHeader: false,
  } as unknown as DocTableRow;
  return {
    colWidths: [50],
    rows: [row],
    borders: emptyBorders(),
    cellMarginTop: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 0,
    cellMarginRight: 0,
    jc: 'left',
  } as unknown as DocTable;
}

// PHYSICAL portrait page 200×300 pt with DISTINCT margins so the logical
// (swapped) frame's coordinates can never coincide with the physical ones:
//   physical top=20 right=30 bottom=40 left=24
//   ⇒ logical  left=20 top=30   right=40  bottom=24 (verticalLayoutSection)
// Flow starts at logical y = 30; the column axis starts at logical x = 20
// (= the physical top margin). Flow budget = 200 − 30 − 24 = 146 pt.
const PHYS = {
  pageWidth: 200, pageHeight: 300,
  marginTop: 20, marginRight: 30, marginBottom: 40, marginLeft: 24,
  headerDistance: 0, footerDistance: 0,
  titlePage: false, evenAndOddHeaders: false,
  textDirection: 'tbRl',
} as unknown as SectionProps;

const CSS_W = 200; // physical page width == canvas CSS width at scale 1
const FLOW_TOP = 30; // logical body top (physical right margin)
const COL_TOP = 20; // physical column top (logical marginLeft)
const TABLE_W = 50; // fixed tcW ⇒ physical table width

const CELL_TEXT_1 = 'aaa '.repeat(12).trim();
const CELL_TEXT_2 = 'bbb '.repeat(12).trim();

function verticalTableDoc(): DocxDocumentModel {
  return {
    section: PHYS,
    body: [
      { type: 'table', ...fixedTable(CELL_TEXT_1, 80) },
      { type: 'paragraph', ...bodyParagraph('zz') },
      { type: 'table', ...fixedTable(CELL_TEXT_2, null) },
      { type: 'paragraph', ...bodyParagraph('qq') },
    ],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

function documentLayout(doc: DocxDocumentModel): DocumentLayout {
  const measure = makeRecordingCanvas();
  return layoutDocument(doc, createLayoutServices(doc, {
    localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]),
    measureContext: measure.canvas.getContext('2d'),
  }), { currentDateMs: 0 });
}

function documentServices(doc: DocxDocumentModel): ReturnType<typeof createLayoutServices> {
  const measure = makeRecordingCanvas();
  return createLayoutServices(doc, {
    localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]),
    measureContext: measure.canvas.getContext('2d'),
  });
}

function bodyColumn(page: LayoutPage, node: PaintNode): number {
  const region = page.sectionRegions.find((candidate) => candidate.flowDomainIds.includes(node.flowDomainId));
  return region?.flowDomainIds.indexOf(node.flowDomainId) ?? -1;
}

function floatingTable(text: string, widthPt = 50, mixedAxis = false): BodyElement {
  const table = fixedTable(text, 20);
  table.colWidths = [widthPt];
  table.rows[0]!.cells[0]!.widthPt = widthPt;
  return {
    ...table,
    type: 'table',
    tblpPr: {
      leftFromText: 0,
      rightFromText: 0,
      topFromText: 0,
      bottomFromText: 0,
      horzAnchor: mixedAxis ? 'text' : 'page',
      horzSpecified: true,
      vertAnchor: 'page',
      tblpX: mixedAxis ? 5 : 120,
      tblpY: 20,
    },
    overlap: 'never',
  } as unknown as BodyElement;
}

function verticalNestedFloatingTableDoc(
  relocate = false,
  twoFloats = false,
  mixedAxis = false,
): DocxDocumentModel {
  const outer = fixedTable('', null);
  outer.rows[0]!.cells[0]!.content = [
    floatingTable('FLOAT', twoFloats ? 60 : 50, mixedAxis),
    ...(twoFloats ? [floatingTable('FLOAT2', 60)] : []),
    { type: 'paragraph', ...bodyParagraph('anchor text wraps around nested table') },
  ] as unknown as DocTableCell['content'];
  if (twoFloats) {
    const parentBorder = { style: 'single', width: 1, color: 'ff00ff' };
    outer.borders = {
      top: parentBorder,
      bottom: parentBorder,
      left: parentBorder,
      right: parentBorder,
      insideH: null,
      insideV: null,
    };
  }
  const leading = fixedTable('', 20);
  leading.colWidths = [100];
  leading.rows[0]!.cells[0]!.widthPt = 100;
  return {
    section: PHYS,
    body: [
      ...(relocate ? [{ type: 'table', ...leading }] : []),
      { type: 'table', ...outer },
    ],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

async function renderRuns(): Promise<{ runs: DocxTextRunInfo[]; rectCalls: RectCall[] }> {
  const { canvas, rectCalls } = makeRecordingCanvas();
  const runs: DocxTextRunInfo[] = [];
  await renderDocumentToCanvas(verticalTableDoc(), canvas, 0, {
    dpr: 1,
    width: PHYS.pageWidth,
    onTextRun: (r) => runs.push(r),
  });
  return { runs, rectCalls };
}

describe('vertical (tbRl) table cells render upright/horizontal (§17.6.20 + §17.4.80, #988 ④)', () => {
  it('cell text is horizontal at the fixed physical width, from the physical column top', async () => {
    const { runs } = await renderRuns();
    const cellRuns = runs.filter((r) => r.text.startsWith('aaa'));
    expect(cellRuns.length).toBeGreaterThanOrEqual(2);
    for (const r of cellRuns) {
      // Horizontal: no +90° overlay rotation.
      expect(r.transform, `run ${JSON.stringify(r)}`).toBeUndefined();
      // Fixed tcW ⇒ the physical x band [cssW − flowTop − tableW, +tableW].
      expect(r.x).toBeGreaterThanOrEqual(CSS_W - FLOW_TOP - TABLE_W - 2);
      expect(r.x + r.w).toBeLessThanOrEqual(CSS_W - FLOW_TOP + 2);
    }
    // Lines share the same x (left-aligned) and stack DOWN the physical page
    // from the physical column top (top content margin).
    const ys = cellRuns.map((r) => r.y);
    expect(new Set(cellRuns.map((r) => Math.round(r.x))).size).toBe(1);
    expect(Math.min(...ys)).toBeGreaterThanOrEqual(COL_TOP - 1);
    expect(Math.min(...ys)).toBeLessThanOrEqual(COL_TOP + 15);
    for (let i = 1; i < ys.length; i++) expect(ys[i]).toBeGreaterThan(ys[i - 1]);
  });

  it('trHeight exact clips at the physical row height; auto grows past it', async () => {
    const { runs, rectCalls } = await renderRuns();
    // §17.4.80 exact ⇒ an 80 pt clip band. The retained painter submits
    // point-space clip geometry before the canvas placement transform, while
    // the text-run assertions below observe the resulting physical placement.
    const clipRect = rectCalls.find(
      (r) => Math.abs(r.h - 80) < 1e-6,
    );
    expect(clipRect, 'exact row must clip at its physical Y band').toBeDefined();
    // auto ⇒ the row grows: content extends well past the 80 pt exact height.
    const autoRuns = runs.filter((r) => r.text.startsWith('bbb'));
    expect(autoRuns.length).toBeGreaterThanOrEqual(2);
    for (const r of autoRuns) expect(r.transform).toBeUndefined();
    expect(Math.max(...autoRuns.map((r) => r.y))).toBeGreaterThan(COL_TOP + 80);
  });

  it('the table advances the flow by its PHYSICAL WIDTH; body text stays vertical', async () => {
    const { runs } = await renderRuns();
    const afterExact = runs.find((r) => r.text === 'zz');
    expect(afterExact).toBeDefined();
    // Body text on a vertical page keeps the rotated overlay placement.
    expect(afterExact!.transform).toBe('rotate(90deg)');
    // The paragraph after the exact table starts one TABLE-WIDTH (50, the
    // physical width) past the flow top — NOT one logical-row-height (80) past:
    // place.left = cssW − (flowTop + tableW) = 200 − 80 = 120.
    expect(afterExact!.x).toBeCloseTo(CSS_W - (FLOW_TOP + TABLE_W), 1);
    // And the following auto table + paragraph keep flowing right→left.
    const afterAuto = runs.find((r) => r.text === 'qq');
    expect(afterAuto).toBeDefined();
    expect(afterAuto!.transform).toBe('rotate(90deg)');
    expect(afterAuto!.x).toBeLessThan(afterExact!.x);
  });

  it('pagination charges the physical table width as the flow footprint', () => {
    const { canvas } = makeRecordingCanvas();
    const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
    // Logical flow budget = physical pageWidth − logical top/bottom insets
    // (30 + 24) = 146. A leading ~12 pt line + a 140 pt-EXACT-row table:
    // charging the logical row height (140) overflows onto page 2, while the
    // upright footprint (tableW = 50) fits page 1 beside it.
    const body = [
      { type: 'paragraph', ...bodyParagraph('zz') },
      { type: 'table', ...fixedTable(CELL_TEXT_1, 140) },
    ] as unknown as BodyElement[];
    const layout = layoutBodyModel(body, PHYS, ctx, { 'Times New Roman': 'roman' });
    expect(layout.pages.length).toBe(1);
    expect(layout.pages[0].layers.body.length).toBe(2);
  });

  it('retains upright physical geometry while the placement owns the vertical flow footprint', async () => {
    const doc = verticalTableDoc();
    const layout = documentLayout(doc);
    const tables = layout.pages[0].layers.body.filter((node) => node.kind === 'table');
    expect(tables).toHaveLength(2);

    const [exact, auto] = tables;
    if (exact?.kind !== 'table' || auto?.kind !== 'table') {
      throw new Error('expected retained upright table layouts');
    }

    // §17.6.20: the upright table's rows remain in physical coordinates, so
    // the exact row owns an 80 pt physical row stack. The surrounding vertical
    // story advances by the table's 50 pt physical width instead.
    expect(exact.advancePt).toBeCloseTo(80, 6);
    expect(auto.advancePt).toBeGreaterThan(80);
    expect(exact.flowBounds.heightPt).toBeCloseTo(TABLE_W, 6);
    expect(auto.flowBounds.heightPt).toBeCloseTo(TABLE_W, 6);

    const paint = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, paint.canvas, 0, {
      dpr: 1,
      width: PHYS.pageWidth,
      layoutServices: documentServices(doc),
    });
    expect(paint.measureCalls()).toBe(0);
  });

  it('retains one physical nested-float box for upright wrap, pagination, and paint', async () => {
    const doc = verticalNestedFloatingTableDoc();
    const layout = documentLayout(doc);
    const table = layout.pages[0]!.layers.body.find((node) => node.kind === 'table');
    if (table?.kind !== 'table' || !('resolvedFloatingTables' in table)) {
      throw new Error('expected an upright retained table fragment with nested floats');
    }
    const fragment = table as TableFragmentLayout;
    const nested = fragment.resolvedFloatingTables[0];

    expect(layout.pages).toHaveLength(1);
    expect(nested).toMatchObject({
      bounds: { xPt: 120, yPt: 20, widthPt: 50, heightPt: 20 },
      exclusionBounds: { xPt: 120, yPt: 20, widthPt: 50, heightPt: 20 },
    });
    expect(nested?.source.physicalPageIndex).toBe(0);
    const anchorBlock = fragment.rows[0]?.cells[0]?.blocks.find(
      (block) => block.layout.kind === 'paragraph',
    );
    expect(anchorBlock?.layout.kind === 'paragraph'
      ? anchorBlock.layout.exclusions
      : []).not.toEqual([]);

    const paint = makeRecordingCanvas();
    const runs: DocxTextRunInfo[] = [];
    await renderDocumentToCanvas(doc, paint.canvas, 0, {
      dpr: 1,
      width: PHYS.pageWidth,
      layoutServices: documentServices(doc),
      onTextRun: (run) => runs.push(run),
    });
    expect(paint.measureCalls()).toBe(0);
    expect(runs.find((run) => run.text === 'FLOAT')).toMatchObject({
      x: expect.closeTo(120, 0),
      y: expect.closeTo(20, 0),
    });
  });

  it('finalizes an upright nested float only after relocating to its destination page', async () => {
    const doc = verticalNestedFloatingTableDoc(true);
    const layout = documentLayout(doc);
    const target = layout.pages[1]?.layers.body.find((node) => node.kind === 'table');
    if (target?.kind !== 'table' || !('resolvedFloatingTables' in target)) {
      throw new Error('expected relocated upright retained nested float');
    }
    const fragment = target as TableFragmentLayout;
    const nested = fragment.resolvedFloatingTables[0]!;
    const anchorBlock = fragment.rows[0]?.cells[0]?.blocks.find(
      (block) => block.layout.kind === 'paragraph',
    );
    const expectedLocalExclusion = {
      xPt: nested.exclusionBounds.xPt - nested.source.anchorBounds.xPt,
      yPt: nested.exclusionBounds.yPt - nested.source.anchorBounds.yPt,
      widthPt: nested.exclusionBounds.widthPt,
      heightPt: nested.exclusionBounds.heightPt,
    };

    expect(layout.pages).toHaveLength(2);
    expect(fragment.resolvedFloatingTableCoordinateSpace).toBe('upright-physical-page-points');
    expect(nested.source.physicalPageIndex).toBe(1);
    expect(anchorBlock?.layout.kind === 'paragraph'
      ? anchorBlock.layout.exclusions
      : []).toContainEqual(expect.objectContaining({ bounds: expectedLocalExclusion }));

    const paint = makeRecordingCanvas();
    const runs: DocxTextRunInfo[] = [];
    await renderDocumentToCanvas(doc, paint.canvas, 1, {
      dpr: 1,
      width: PHYS.pageWidth,
      layoutServices: documentServices(doc),
      onTextRun: (run) => runs.push(run),
    });
    expect(paint.measureCalls()).toBe(0);
    expect(runs.find((run) => run.text === 'FLOAT')).toMatchObject({
      x: expect.closeTo(nested.bounds.xPt, 0),
      y: expect.closeTo(nested.bounds.yPt, 0),
    });
  });

  it('resolves two upright nested floats in source order and paints each before the parent border', async () => {
    const doc = verticalNestedFloatingTableDoc(false, true);
    const layout = documentLayout(doc);
    const table = layout.pages[0]!.layers.body.find((node) => node.kind === 'table');
    if (table?.kind !== 'table' || !('resolvedFloatingTables' in table)) {
      throw new Error('expected two retained upright nested floats');
    }
    const fragment = table as TableFragmentLayout;
    const [first, second] = fragment.resolvedFloatingTables;

    expect(fragment.resolvedFloatingTableCoordinateSpace).toBe('upright-physical-page-points');
    expect(fragment.resolvedFloatingTables).toHaveLength(2);
    expect(first?.bounds).toEqual({ xPt: 120, yPt: 20, widthPt: 60, heightPt: 20 });
    expect(second?.bounds).toEqual({ xPt: 120, yPt: 40, widthPt: 60, heightPt: 20 });

    const paint = makeRecordingCanvas();
    const runs: DocxTextRunInfo[] = [];
    await renderDocumentToCanvas(doc, paint.canvas, 0, {
      dpr: 1,
      width: PHYS.pageWidth,
      layoutServices: documentServices(doc),
      onTextRun: (run) => runs.push(run),
    });
    expect(paint.measureCalls()).toBe(0);
    expect(runs.filter((run) => run.text === 'FLOAT' || run.text === 'FLOAT2')
      .map((run) => run.text)).toEqual(['FLOAT', 'FLOAT2']);
    const parentBorderIndex = paint.paintEvents.findIndex(
      (event) => event.kind === 'stroke' && event.color === '#ff00ff',
    );
    expect(parentBorderIndex).toBeGreaterThanOrEqual(0);
  });

  it('composes the upright cell-column X axis with the physical page Y axis', () => {
    const doc = verticalNestedFloatingTableDoc(false, false, true);
    const layout = documentLayout(doc);
    const table = layout.pages[0]!.layers.body.find((node) => node.kind === 'table');
    if (table?.kind !== 'table' || !('resolvedFloatingTables' in table)) {
      throw new Error('expected a mixed-axis upright nested float');
    }
    const [placement] = (table as TableFragmentLayout).resolvedFloatingTables;

    if (!placement?.source.columnBounds) throw new Error('expected the retained cell column');
    // Horizontal text placement composes the retained physical cell-column X
    // with tblpX, while the independent page-relative vertical axis stays y=20.
    expect(placement.bounds).toEqual({
      xPt: placement.source.columnBounds.xPt + 5,
      yPt: 20,
      widthPt: 50,
      heightPt: 20,
    });
  });

  it('records the destination column on an upright retained placement', () => {
    const wide = fixedTable('', 20);
    wide.colWidths = [140];
    wide.rows[0].cells[0].widthPt = 140;
    const narrow = fixedTable('', 20);
    const physical = {
      ...PHYS,
      columns: { count: 2, spacePt: 10, equalWidth: true, sep: false, cols: [] },
    } as SectionProps;
    const body = [
      { type: 'table', ...wide },
      { type: 'table', ...narrow },
    ] as unknown as BodyElement[];
    const measure = makeRecordingCanvas();
    const layout = layoutBodyModel(
      body,
      physical,
      measure.canvas.getContext('2d') as CanvasRenderingContext2D,
      { 'Times New Roman': 'roman' },
    );
    const page = layout.pages[0];
    const tables = page.layers.body.filter((node) => node.kind === 'table');
    expect(tables).toHaveLength(2);
    expect(bodyColumn(page, tables[0])).toBe(0);
    expect(bodyColumn(page, tables[1])).toBe(1);
  });
});
