import { describe, it, expect } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { paintLayoutPage } from './paint/canvas-page.js';
import type { DocumentLayout } from './layout/types.js';
import { testFontSnapshot } from './layout/test-font-snapshot.js';
import type { BodyElement, DocParagraph, DocTable, DocTableRow, DocTableCell, DocxDocumentModel, SectionProps } from './types';

// ─────────────────────────────────────────────────────────────────────────────
// A5 P4 — retained TABLE-CELL paragraph layout.
//
// Pagination acquires each cell paragraph once in scale-1 point geometry and keeps
// it inside TableLayout/TableFragmentLayout. Paint consumes that retained tree and
// never writes line caches onto the parsed document model.
//
// These tests pin, using the SAME cross-context flow the public renderPage uses
// (paginate ctx from OffscreenCanvas(1,1) ≠ paint ctx), that:
//   (a) cell paragraph paint makes zero measureText calls;
//   (b) retained acquisition does not create obsolete parser-side line stamps;
//   (c) the wrap PARTITION is zoom-invariant: the same table painted at scale 1
//       and at scale 0.75 breaks each cell paragraph at the same points (line
//       count + relative segment order);
//   (d) nested tables retain their inner cell paragraph geometry too.
// ─────────────────────────────────────────────────────────────────────────────

interface Call { op: 'fill' | 'stroke' | 'img'; text: string; x: number; y: number; font: string; }

/** Linear-metric measure stub for the paginate ctx (node lacks OffscreenCanvas).
 *  Same advance law as the recording paint ctx so paginate == paint measures. */
function makeMeasureStubCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      const per = p * 0.5;
      return {
        width: [...s].length * per,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}
(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
  getContext() { return makeMeasureStubCtx(); }
};

/** A recording paint ctx with a font-size-LINEAR advance. Records every
 *  text/image draw and counts forbidden paint-time measureText calls. */
function makeRecordingCanvas(): { canvas: HTMLCanvasElement; calls: Call[]; measures: () => number } {
  let font = '10px serif';
  const calls: Call[] = [];
  let measures = 0;
  let transform = { scaleX: 1, scaleY: 1, translateX: 0, translateY: 0 };
  const stack: typeof transform[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      measures++;
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      const per = p * 0.5;
      return {
        width: [...s].length * per,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() { stack.push({ ...transform }); },
    restore() { transform = stack.pop() ?? transform; },
    setTransform(a: number, _b: number, _c: number, d: number, e: number, f: number) {
      transform = { scaleX: a, scaleY: d, translateX: e, translateY: f };
    },
    beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {},
    scale(x: number, y: number) {
      transform.scaleX *= x;
      transform.scaleY *= y;
    },
    translate(x: number, y: number) {
      transform.translateX += transform.scaleX * x;
      transform.translateY += transform.scaleY * y;
    },
    rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage(_img: unknown, x: number, y: number) {
      calls.push({
        op: 'img', text: '',
        x: transform.translateX + transform.scaleX * x,
        y: transform.translateY + transform.scaleY * y,
        font,
      });
    },
    fillText(s: string, x: number, y: number) {
      calls.push({
        op: 'fill', text: s,
        x: transform.translateX + transform.scaleX * x,
        y: transform.translateY + transform.scaleY * y,
        font,
      });
    },
    strokeText(s: string, x: number, y: number) {
      calls.push({
        op: 'stroke', text: s,
        x: transform.translateX + transform.scaleX * x,
        y: transform.translateY + transform.scaleY * y,
        font,
      });
    },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls, measures: () => measures };
}

// ---- Model builders (mirror table-layout-reuse.test / layout-lines-reuse) ------
function textRun(text: string): DocParagraph['runs'][number] {
  return {
    type: 'text', text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: '',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  } as DocParagraph['runs'][number];
}
function para(text: string, over: Partial<DocParagraph> = {}): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [textRun(text)],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
    ...over,
  } as unknown as DocParagraph;
}
function emptyBorders() {
  return { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };
}
function cellOf(content: DocParagraph[] | DocTable[], widthPt = 120): DocTableCell {
  return {
    content: content.map((c) => ((c as { type?: string }).type === 'table' ? c : { type: 'paragraph', ...(c as DocParagraph) })),
    colSpan: 1, vMerge: null, borders: emptyBorders(), background: null, vAlign: 'top', widthPt,
  } as unknown as DocTableCell;
}
function cell(text: string, widthPt = 120): DocTableCell { return cellOf([para(text)], widthPt); }
function row(cells: DocTableCell[]): DocTableRow {
  return { cells, rowHeight: null, rowHeightRule: 'auto', isHeader: false } as unknown as DocTableRow;
}
function tableOf(rows: DocTableRow[], colWidths: number[]): DocTable {
  return {
    type: 'table', colWidths, rows, borders: emptyBorders(),
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 2, cellMarginRight: 2,
    jc: 'left', layout: 'fixed',
  } as unknown as DocTable;
}
/** A table with `nRows` rows × 2 cols of multi-word wrapping content. */
function wrapTable(nRows: number): DocTable {
  const rows: DocTableRow[] = [];
  for (let r = 0; r < nRows; r++) {
    rows.push(row([
      cell(`row ${r} left ` + Array.from({ length: 7 }, (_, i) => `wa${i}`).join(' '), 120),
      cell(`row ${r} right ` + Array.from({ length: 7 }, (_, i) => `wb${i}`).join(' '), 160),
    ]));
  }
  return tableOf(rows, [120, 160]);
}
function doc(body: BodyElement[], pageHeight = 200): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 300, pageHeight,
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

/** Render every page at `width` (paint scale = width / pageWidth), returning the
 *  concatenated paint stream per page + the total paint-time measureText count. */
async function renderAll(layout: DocumentLayout, width: number): Promise<{ perPage: Call[][]; measures: number }> {
  const perPage: Call[][] = [];
  let measures = 0;
  for (let p = 0; p < layout.pages.length; p++) {
    const rec = makeRecordingCanvas();
    await paintLayoutPage(layout, p, rec.canvas, {
      dpr: 1,
      scale: width / layout.pages[p].geometry.widthPt,
    });
    perPage.push(rec.calls);
    measures += rec.measures();
  }
  return { perPage, measures };
}

async function retainedPaint(model: DocxDocumentModel, width = 300): Promise<{
  layout: DocumentLayout; drawn: number; measures: number; streams: Call[][];
}> {
  const layout = layoutDocument(model, createLayoutServices(model, {
    localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]),
  }), { currentDateMs: 0 });
  const painted = await renderAll(layout, width);
  return {
    layout,
    drawn: painted.perPage.flat().filter((call) => call.op !== 'img').length,
    measures: painted.measures,
    streams: painted.perPage,
  };
}

function tableMeasurementGeometry() {
  const model = doc([wrapTable(8) as unknown as BodyElement]);
  const layout = layoutDocument(model, createLayoutServices(model, {
    localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]),
  }), { currentDateMs: 0 });
  const fragments = layout.pages.flatMap((page) => page.layers.body.filter((node) => node.kind === 'table'));
  return {
    pageCount: layout.pages.length,
    rowHeightsPt: fragments.flatMap((fragment) => fragment.rows.map((tableRow) => tableRow.heightPt)),
    cellLineCounts: fragments.flatMap((fragment) => fragment.rows.flatMap((tableRow) =>
      tableRow.cells.flatMap((tableCell) => tableCell.blocks.map((block) =>
        block.layout.kind === 'paragraph' ? block.layout.lines.length : null,
      )))),
  };
}

describe('table-cell paragraph line reuse — B2 T2', () => {
  it('preserves table-cell heights and reusable line counts', () => {
    const geometry = tableMeasurementGeometry();

    expect(geometry.pageCount).toBe(2);
    expect(geometry.rowHeightsPt).toEqual([
      22.998046875,
      22.998046875,
      22.998046875,
      22.998046875,
      22.998046875,
      22.998046875,
      22.998046875,
      11.4990234375,
      11.4990234375,
    ]);
    expect(geometry.cellLineCounts).toEqual([
      ...Array.from({ length: 14 }, () => 2),
      ...Array.from({ length: 4 }, () => 1),
    ]);
  });

  it('(a) retained cell paragraphs paint without measuring', async () => {
    const r = await retainedPaint(doc([wrapTable(16) as unknown as BodyElement]));
    expect(r.layout.pages.length).toBeGreaterThan(1); // table split across pages
    expect(r.drawn).toBeGreaterThan(0); // really painted cell text
    expect(r.measures).toBe(0);
  });

  it('(b) retained cell acquisition never writes obsolete parser line stamps', async () => {
    const model = doc([wrapTable(16) as unknown as BodyElement]);
    const layout = layoutDocument(model, createLayoutServices(model, {
      localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]),
    }), { currentDateMs: 0 });
    const before = await renderAll(layout, 300);
    for (const element of model.body) {
      if (element.type !== 'table') continue;
      for (const rw of element.rows) for (const c of rw.cells) for (const ce of c.content) {
        expect(ce).not.toHaveProperty('layoutLines');
        expect(ce).not.toHaveProperty('layoutLinesInputs');
      }
    }
    const after = await renderAll(layout, 300);
    expect(before.measures).toBe(0);
    expect(after.measures).toBe(0);
    expect(after.perPage).toEqual(before.perPage);
  });

  it('(c) wrap partition is zoom-invariant: scale 1 and scale 0.75 break each cell paragraph identically', async () => {
    const model = doc([wrapTable(10) as unknown as BodyElement]);
    const layout = layoutDocument(model, createLayoutServices(model, { localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]) }), { currentDateMs: 0 });
    // Same paint text at two scales; the fillText SEQUENCE (line partition) must be
    // identical — only the x/y coordinates scale. Compare the per-line text runs.
    const at1 = await renderAll(layout, 300);   // scale 1
    const at075 = await renderAll(layout, 225);  // scale 0.75 (225 = 300*0.75)
    expect(at1.perPage.length).toBe(at075.perPage.length);
    for (let p = 0; p < at1.perPage.length; p++) {
      const text1 = at1.perPage[p].filter((c) => c.op !== 'img').map((c) => c.text);
      const text075 = at075.perPage[p].filter((c) => c.op !== 'img').map((c) => c.text);
      // Identical drawn-text sequence ⇒ identical wrap partition at both zooms.
      expect(text075).toEqual(text1);
    }
    // And the x positions scale by 0.75 (partition reused, geometry rehydrated).
    const fills1 = at1.perPage[0].filter((c) => c.op === 'fill');
    const fills075 = at075.perPage[0].filter((c) => c.op === 'fill');
    expect(fills075.length).toBe(fills1.length);
    for (let i = 0; i < fills1.length; i++) {
      expect(fills075[i].x).toBeCloseTo(fills1[i].x * 0.75, 6);
    }
  });

  it('(d) nested table: inner-cell paragraphs are retained and painted measure-free', async () => {
    // Outer cell contains a nested table whose own cells wrap. The nested table's
    // cell paragraphs are acquired at scale 1 through the same retained recursion.
    const inner = tableOf([
      row([cell('inner ' + Array.from({ length: 6 }, (_, i) => `x${i}`).join(' '), 100)]),
      row([cell('more ' + Array.from({ length: 6 }, (_, i) => `y${i}`).join(' '), 100)]),
    ], [100]);
    const outer = tableOf([
      row([cellOf([inner] as unknown as DocTable[], 140), cell('side ' + Array.from({ length: 6 }, (_, i) => `z${i}`).join(' '), 140)]),
    ], [140, 140]);
    const r = await retainedPaint(doc([outer as unknown as BodyElement]));
    expect(r.drawn).toBeGreaterThan(0);
    expect(r.measures).toBe(0);
    // The inner table's text was actually drawn (nested content painted).
    const drewInner = r.streams.some((page) => page.some((c) => c.text.startsWith('inner') || c.text.startsWith('x')));
    expect(drewInner).toBe(true);
  });

  it('(e) same page painted twice is identical (the retained tree is not mutated)', async () => {
    const model = doc([wrapTable(16) as unknown as BodyElement]);
    const layout = layoutDocument(model, createLayoutServices(model, { localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]) }), { currentDateMs: 0 });
    const first = await renderAll(layout, 300);
    const second = await renderAll(layout, 300);
    for (let p = 0; p < first.perPage.length; p++) expect(second.perPage[p]).toEqual(first.perPage[p]);
  });
});
