import { describe, it, expect } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { paintLayoutPage } from './paint/canvas-page.js';
import type { DocumentLayout, LayoutPage, PaintNode } from './layout/types.js';
import { testFontSnapshot } from './layout/test-font-snapshot.js';
import { stableFingerprint } from './layout/fingerprint.js';
import type { TableFragmentLayout } from './layout/table-pagination.js';
import type { FlowFragment } from './layout/flow-fragment.js';
import type {
  BodyElement,
  DocParagraph,
  DocTable,
  DocxDocumentModel,
  SectionProps,
  } from './types';

// ─────────────────────────────────────────────────────────────────────────────
// A3 retained paragraph acquisition and paint invariants.
//
// Every body/cell paragraph is acquired once as immutable point-space geometry.
// Production paint consumes that node without measuring or mutating it. Tests below
// assert semantic retained geometry and deterministic device mapping; the removed
// paragraph fallback/toggle paths are deliberately not test oracles.
//
// The retained layout is built once and passed directly to canonical page paint.
// The render width equals the page
// width so the paint scale is exactly 1 (fragment rescale is then a no-op → a
// migrated paragraph paints with zero measureText calls).
// ─────────────────────────────────────────────────────────────────────────────

interface Call {
  op: 'fill' | 'stroke' | 'img';
  text: string;
  x: number;
  y: number;
  font: string;
  scaleX: number;
  scaleY: number;
}

/** Canonical layout builds its measure ctx from `new OffscreenCanvas(1,1)`,
 *  which the node test env lacks. Polyfill it with the SAME linear metric the
 *  recording paint canvas uses (glyph width = fontPx·0.5) so the paginate ctx and
 *  the paint ctx measure identically — the condition under which Stage 1
 *  guarantees byte-identical reuse. (Real cross-context metric parity is a browser
 *  Canvas invariant, exercised end-to-end by `pnpm vrt`.) */
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

/** A recording context whose glyph advance is LINEAR in the font px size, so the
 *  scale-1 paginate lines and the scale-1 paint lines are bit-identical (this
 *  test is about the reuse mechanism, not font hinting). It records every text
 *  and image draw with position + font for an exact stream comparison. */
function makeRecordingCanvas(): { canvas: HTMLCanvasElement; calls: Call[]; measures: () => number } {
  let font = '10px serif';
  const calls: Call[] = [];
  let measures = 0;
  let transform = { scaleX: 1, scaleY: 1, translateX: 0, translateY: 0 };
  const transformStack: typeof transform[] = [];
  const canvas: {
    width: number;
    height: number;
    style: Record<string, string>;
    getContext?: () => unknown;
  } = { width: 0, height: 0, style: {} };
  const record = (op: Call['op'], text: string, x: number, y: number): void => {
    calls.push({
      op,
      text,
      x: transform.translateX + transform.scaleX * x,
      y: transform.translateY + transform.scaleY * y,
      font,
      scaleX: transform.scaleX,
      scaleY: transform.scaleY,
    });
  };
  const ctx = {
    canvas,
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
    save() { transformStack.push({ ...transform }); },
    restore() { transform = transformStack.pop() ?? transform; },
    setTransform(a: number, _b: number, _c: number, d: number, e: number, f: number) {
      transform = { scaleX: a, scaleY: d, translateX: e, translateY: f };
    },
    beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {},
    scale(sx: number, sy: number) {
      transform.scaleX *= sx;
      transform.scaleY *= sy;
    },
    translate(x: number, y: number) {
      transform.translateX += transform.scaleX * x;
      transform.translateY += transform.scaleY * y;
    },
    rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage(_img: unknown, ...args: number[]) {
      const x = args.length >= 8 ? args[4] : args[0];
      const y = args.length >= 8 ? args[5] : args[1];
      record('img', '', x ?? 0, y ?? 0);
    },
    fillText(s: string, x: number, y: number) { record('fill', s, x, y); },
    strokeText(s: string, x: number, y: number) { record('stroke', s, x, y); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  canvas.getContext = () => ctx;
  return { canvas: canvas as unknown as HTMLCanvasElement, calls, measures: () => measures };
}

function para(text: string, over: Partial<DocParagraph> = {}): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'Times New Roman',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
    } as DocParagraph['runs'][number]],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
    ...over,
  } as unknown as DocParagraph;
}

function doc(body: BodyElement[], pageHeight = 60, settings?: Record<string, unknown>): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 200, pageHeight,
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
    ...(settings ? { settings } : {}),
  } as unknown as DocxDocumentModel;
}

/** Render every canonical page at scale 1, returning its paint stream. */
async function renderAllPages(layout: DocumentLayout): Promise<{ perPage: Call[][]; measures: number }> {
  const perPage: Call[][] = [];
  let measures = 0;
  for (let p = 0; p < layout.pages.length; p++) {
    const rec = makeRecordingCanvas();
    await paintLayoutPage(layout, p, rec.canvas, { dpr: 1, scale: 1 }, { paint() {} });
    perPage.push(rec.calls);
    measures += rec.measures();
  }
  return { perPage, measures };
}

function retainedFlowGeometry(fragment: FlowFragment | PaintNode): unknown {
  if (fragment.kind === 'paragraph') return fragment;
  if (fragment.kind === 'table') {
    return {
      kind: fragment.kind,
      columnWidthsPt: fragment.columnWidthsPt,
      rows: fragment.rows.map((row) => ({
        heightPt: row.flowBounds.heightPt,
        repeatedHeader: row.repeatedHeader ?? false,
        cells: row.cells.map((cell) => ({
          verticalMerge: cell.verticalMerge,
          boxHeightPt: cell.flowBounds.heightPt,
          blocks: cell.blocks.map((block) => retainedFlowGeometry(block.layout)),
        })),
      })),
    };
  }
  return fragment;
}

function columnIndex(page: LayoutPage, node: PaintNode): number {
  const region = page.sectionRegions.find((candidate) => candidate.flowDomainIds.includes(node.flowDomainId));
  return region?.flowDomainIds.indexOf(node.flowDomainId) ?? -1;
}

function retainedFingerprints(layout: DocumentLayout): string[] {
  return layout.pages.flatMap((page) => page.layers.body.map((node) =>
    stableFingerprint('retained-flow', {
      columnIndex: columnIndex(page, node),
      bounds: node.flowBounds,
      advancePt: node.advancePt,
      fragment: retainedFlowGeometry(node),
    })));
}

async function renderRetainedAtScale(
  layout: DocumentLayout,
  scale: 1 | 2,
): Promise<{ calls: Call[][]; measures: number }> {
  const calls: Call[][] = [];
  let measures = 0;
  for (let pageIndex = 0; pageIndex < layout.pages.length; pageIndex++) {
    const rec = makeRecordingCanvas();
    await paintLayoutPage(layout, pageIndex, rec.canvas, {
      dpr: 1,
      scale,
    }, { paint() {} });
    calls.push(rec.calls);
    measures += rec.measures();
  }
  return { calls, measures };
}

/** Final A3 invariant: paint consumes immutable point-space geometry. It may only
 *  apply the page's device transform; it may neither reshape nor mutate the node. */
async function assertRetainedScaleInvariant(
  model: DocxDocumentModel,
): Promise<{ layout: DocumentLayout; fingerprints: string[] }> {
  const layout = layoutDocument(model, createLayoutServices(model, {
    localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]),
  }), { currentDateMs: 0 });
  const before = retainedFingerprints(layout);
  expect(before.length).toBeGreaterThan(0);

  const at1 = await renderRetainedAtScale(layout, 1);
  expect(at1.measures).toBe(0);
  expect(retainedFingerprints(layout)).toEqual(before);

  const at2 = await renderRetainedAtScale(layout, 2);
  expect(at2.measures).toBe(0);
  expect(retainedFingerprints(layout)).toEqual(before);
  expect(at2.calls).toHaveLength(at1.calls.length);

  for (let pageIndex = 0; pageIndex < at1.calls.length; pageIndex++) {
    expect(at2.calls[pageIndex]).toHaveLength(at1.calls[pageIndex].length);
    for (let callIndex = 0; callIndex < at1.calls[pageIndex].length; callIndex++) {
      const point = at1.calls[pageIndex][callIndex];
      const device = at2.calls[pageIndex][callIndex];
      expect({ op: device.op, text: device.text, font: device.font }).toEqual({
        op: point.op, text: point.text, font: point.font,
      });
      expect(device.x).toBe(point.x * 2);
      expect(device.y).toBe(point.y * 2);
      expect(device.scaleX).toBe(point.scaleX * 2);
      expect(device.scaleY).toBe(point.scaleY * 2);
    }
  }
  return { layout, fingerprints: before };
}

/** Assert production consumes the retained result deterministically. */
async function assertRetainedPaint(model: DocxDocumentModel): Promise<{
  pages: number;
  drawn: number;
  split: boolean;
  measures: number;
  streams: Call[][];
}> {
  const layout = layoutDocument(model, createLayoutServices(model, { localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]) }), { currentDateMs: 0 });
  const split = layout.pages.some((page) => page.layers.body.some((node) => node.kind === 'paragraph' && node.continuation?.continuesOnNext));
  const before = retainedFingerprints(layout);
  const first = await renderAllPages(layout);
  expect(retainedFingerprints(layout)).toEqual(before);
  const second = await renderAllPages(layout);
  expect(second.perPage).toEqual(first.perPage);
  expect(retainedFingerprints(layout)).toEqual(before);
  const drawn = first.perPage.flat().filter((call) => call.op !== 'img').length;
  return {
    pages: layout.pages.length,
    drawn,
    split,
    measures: first.measures + second.measures,
    streams: first.perPage,
  };
}

describe('body retained paragraph acquisition and paint', () => {
  it('paints a split single-column paragraph without measuring or mutation', async () => {
    const text = Array.from({ length: 120 }, () => 'w').join(' ');
    const r = await assertRetainedPaint(doc([para(text) as unknown as BodyElement]));
    expect(r.pages).toBeGreaterThan(1); // really split
    expect(r.split).toBe(true);         // continuation slices present
    expect(r.drawn).toBeGreaterThan(0); // really painted
    expect(r.measures).toBe(0);
  });

  it('justified paragraph (both): slack distribution over fragment segments is identical', async () => {
    const text = Array.from({ length: 120 }, (_, i) => (i % 3 === 0 ? 'lorem' : 'ipsum')).join(' ');
    const r = await assertRetainedPaint(doc([para(text, { alignment: 'both' }) as unknown as BodyElement]));
    expect(r.pages).toBeGreaterThan(1);
    expect(r.drawn).toBeGreaterThan(0);
    expect(r.measures).toBe(0);
  });

  it('paints a retained CJK per-glyph partition without measuring', async () => {
    const text = 'あ'.repeat(200);
    const r = await assertRetainedPaint(doc([para(text) as unknown as BodyElement]));
    expect(r.pages).toBeGreaterThan(1);
    expect(r.drawn).toBeGreaterThan(0);
    expect(r.measures).toBe(0);
  });

  it('numbered list that splits acquires the marker-aware body partition once', async () => {
    const numbering = { numId: 1, level: 0, format: 'decimal', text: '1.',
      indentLeft: 36, tab: 36, suff: 'tab', jc: 'left' } as unknown as DocParagraph['numbering'];
    const text = Array.from({ length: 120 }, () => 'w').join(' ');
    const p = para(text, {
      numbering, indentLeft: 36, indentFirst: -18,
    });
    const r = await assertRetainedPaint(doc([p as unknown as BodyElement]));
    expect(r.pages).toBeGreaterThan(1);
    expect(r.drawn).toBeGreaterThan(0);
    expect(r.measures).toBe(0);
  });

  it('NUMPAGES converges before retained acquisition and paints without measuring', async () => {
    const text = Array.from({ length: 120 }, () => 'w').join(' ');
    const p = para(text);
    (p.runs as unknown[]).push({
      type: 'field', fieldType: 'numPages', instruction: 'NUMPAGES', fallbackText: '?',
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', background: null,
    });
    const r = await assertRetainedPaint(doc([p as unknown as BodyElement]));
    expect(r.pages).toBeGreaterThan(1);
    expect(r.split).toBe(true);
    expect(r.measures).toBe(0);
    const drewTotal = r.streams.some((page) => page.some((c) => c.text === String(r.pages)));
    expect(drewTotal).toBe(true);
  });

  it('PAGE is acquired from the destination page context', async () => {
    const pageField = para('');
    (pageField.runs as unknown[]).push({
      type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '?',
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', background: null,
    });
    const model = doc([
      para('first') as unknown as BodyElement,
      { type: 'pageBreak' } as BodyElement,
      pageField as unknown as BodyElement,
    ]);
    const layout = layoutDocument(model, createLayoutServices(model, {
      localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]),
    }), { currentDateMs: 0 });
    expect(layout.pages).toHaveLength(2);
    const production = await renderAllPages(layout);

    expect(production.measures).toBe(0);
    expect(production.perPage[1].some((call) => call.text === '2')).toBe(true);
  });

  it('acquires a PAGE field from the continuation slice that contains its occurrence', async () => {
    const split = para(Array.from({ length: 120 }, () => 'w').join(' '));
    (split.runs as unknown[]).push({
      type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '?',
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', background: null,
    });
    const model = doc([split as unknown as BodyElement]);
    const layout = layoutDocument(model, createLayoutServices(model, {
      localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]),
    }), { currentDateMs: 0 });
    const fieldPageIndex = layout.pages.findIndex((page) => page.layers.body.some((node) =>
      node.kind === 'paragraph'
        && node.lines.some((line) => line.placements.some((placement) =>
          placement.kind === 'text' && placement.dependency === 'page'))));
    expect(fieldPageIndex).toBeGreaterThan(0);

    const production = await renderAllPages(layout);
    expect(production.measures).toBe(0);
    expect(production.perPage[fieldPageIndex].some((call) =>
      call.text === String(fieldPageIndex + 1))).toBe(true);
  });

  it('retains the two-digit PAGE result and its measured advance on page ten', async () => {
    const pageTen = para('');
    (pageTen.runs as unknown[]).push({
      type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '?',
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', background: null,
    });
    const body: BodyElement[] = [];
    for (let pageIndex = 0; pageIndex < 10; pageIndex += 1) {
      if (pageIndex > 0) body.push({ type: 'pageBreak' } as BodyElement);
      body.push((pageIndex === 9 ? pageTen : para(`page ${pageIndex + 1}`)) as unknown as BodyElement);
    }
    const model = doc(body);
    const layout = layoutDocument(model, createLayoutServices(model, {
      localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]),
    }), { currentDateMs: 0 });
    expect(layout.pages).toHaveLength(10);
    const pageTenFragment = layout.pages[9].layers.body
      .filter((node) => node.kind === 'paragraph')
      .find((fragment) => fragment.lines.some((line) => line.placements.some((placement) =>
      placement.kind === 'text' && placement.dependency === 'page')));
    const pagePlacement = pageTenFragment?.lines.flatMap((line) => line.placements)
      .find((placement) => placement.kind === 'text' && placement.dependency === 'page');
    expect(pagePlacement).toMatchObject({ text: '10', advancePt: 10 });

    const production = await renderAllPages(layout);
    expect(production.measures).toBe(0);
    expect(production.perPage[9].some((call) => call.text === '10')).toBe(true);
  });

  it('acquires PAGE with the destination section restart and number format', async () => {
    const pageField = para('');
    (pageField.runs as unknown[]).push({
      type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '?',
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', background: null,
    });
    const model = doc([
      para('first') as unknown as BodyElement,
      { type: 'pageBreak' } as BodyElement,
      pageField as unknown as BodyElement,
    ]);
    model.section = {
      ...model.section,
      pageNumType: { start: 50, fmt: 'upperRoman' },
    };
    const layout = layoutDocument(model, createLayoutServices(model, {
      localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]),
    }), { currentDateMs: 0 });
    const production = await renderAllPages(layout);

    expect(production.measures).toBe(0);
    expect(production.perPage[1].some((call) => call.text === 'LI')).toBe(true);
  });

  it('custom kinsoku settings retain the acquired partition across value-equal rules', async () => {
    // Value-equal kinsoku rules may contain distinct Set instances; retained paint
    // must consume the acquired partition without treating identity as semantics.
    const text = Array.from({ length: 120 }, () => 'w').join(' ');
    const model = doc([para(text) as unknown as BodyElement], 60,
      { kinsoku: true, noLineBreaksBefore: '、。！' , noLineBreaksAfter: '（「' });
    const r = await assertRetainedPaint(model);
    expect(r.pages).toBeGreaterThan(1);
    expect(r.measures).toBe(0);
  });

  it('zoom maps the same retained point geometry through the device transform', async () => {
    const text = Array.from({ length: 120 }, () => 'w').join(' ');
    const model = doc([para(text) as unknown as BodyElement]); // pageHeight 60 → splits
    const { layout } = await assertRetainedScaleInvariant(model);
    expect(layout.pages.length).toBeGreaterThan(1);
    expect(layout.pages.some((page) => page.layers.body.some((node) =>
      node.kind === 'paragraph' && node.continuation?.continuesOnNext))).toBe(true);
  });

  it('retains first-line and hanging indent geometry', async () => {
    const text = Array.from({ length: 120 }, () => 'w').join(' ');
    // First-line indent (positive indentFirst) with left + right indents.
    const firstLine = await assertRetainedPaint(
      doc([para(text, { indentLeft: 24, indentRight: 12, indentFirst: 18 }) as unknown as BodyElement]),
    );
    expect(firstLine.drawn).toBeGreaterThan(0);
    expect(firstLine.measures).toBe(0);
    // Hanging indent (negative first-line) WITHOUT numbering — still fragment-paintable
    // (the numbering exclusion is about numBodyOffset, not a bare hanging indent).
    const hanging = await assertRetainedPaint(
      doc([para(text, { indentLeft: 36, indentFirst: -18 }) as unknown as BodyElement]),
    );
    expect(hanging.drawn).toBeGreaterThan(0);
    expect(hanging.measures).toBe(0);
  });

  it('explicit tab stops retain tab and leader geometry across device scales', async () => {
    // A run with embedded tab characters (\t) and custom tab stops. layoutLines splits
    // on '\t' and advances to the stops (left + right-aligned with a dot leader); the
    // fragment's stored lines carry the tabbed geometry, painted without re-tabbing.
    const cell = Array.from({ length: 6 }, () => 'w').join(' ');
    const tabbed = `${cell}\t${cell}\t${cell}`;
    const tabStops = [
      { pos: 60, alignment: 'left', leader: 'none' },
      { pos: 130, alignment: 'right', leader: 'dot' },
    ] as unknown as DocParagraph['tabStops'];
    const p = para(Array.from({ length: 16 }, () => tabbed).join(' '), { tabStops });
    const { layout } = await assertRetainedScaleInvariant(doc([p as unknown as BodyElement]));
    const tabs = layout.pages.flatMap((page) => page.layers.body.flatMap((node) =>
      node.kind === 'paragraph'
        ? node.lines.flatMap((line) => line.placements.filter((placement) => placement.kind === 'tab'))
        : []
    ));
    expect(tabs.length).toBeGreaterThan(0);
    expect(tabs.some((tab) => tab.leader === 'dot')).toBe(true);
    expect(tabs.every((tab) => tab.advancePt >= 0)).toBe(true);
  });

  it('two-column section retains one point-space layout across device scales', async () => {
    // ECMA-376 §17.6.4 newspaper columns — the body flows through 2 equal columns.
    // Each column-slice fragment records its column band width; paint sets contentW
    // per column, and the placement guard's width check passes (equal columns).
    const columns = { count: 2, spacePt: 12, equalWidth: true, sep: false, cols: [] } as unknown as SectionProps['columns'];
    const text = Array.from({ length: 200 }, () => 'w').join(' ');
    const model = doc([para(text) as unknown as BodyElement]);
    (model.section as SectionProps).columns = columns;
    const { layout } = await assertRetainedScaleInvariant(model);
    const placedColumns = new Set(layout.pages.flatMap((page) =>
      page.layers.body.map((node) => columnIndex(page, node))));
    expect(placedColumns).toEqual(new Set([0, 1]));
  });

  it('same page rendered twice is identical (the shared measured line array is never mutated by paint)', async () => {
    const text = Array.from({ length: 120 }, () => 'w').join(' ');
    const model = doc([para(text) as unknown as BodyElement]);
    const layout = layoutDocument(model, createLayoutServices(model), { currentDateMs: 0 });
    const first = await renderAllPages(layout);
    const second = await renderAllPages(layout);
    for (let p = 0; p < first.perPage.length; p++) expect(second.perPage[p]).toEqual(first.perPage[p]);
  });

  it('acquires placement-specific retained nodes for narrow and wide columns', async () => {
    const text = Array.from({ length: 30 }, () => 'w').join(' ');
    const wideModel = doc([para(text) as unknown as BodyElement], 600); // colW 180
    const narrowModel = doc([para(text) as unknown as BodyElement], 600);
    (narrowModel.section as SectionProps).pageWidth = 120; // colW 100
    const narrowLayout = layoutDocument(narrowModel, createLayoutServices(narrowModel), { currentDateMs: 0 });
    const wideLayout = layoutDocument(wideModel, createLayoutServices(wideModel), { currentDateMs: 0 });
    const narrow = narrowLayout.pages[0].layers.body[0];
    const wide = wideLayout.pages[0].layers.body[0];
    if (narrow?.kind !== 'paragraph' || wide?.kind !== 'paragraph') {
      throw new Error('expected retained paragraph nodes');
    }
    expect(narrow.flowBounds.widthPt).toBe(100);
    expect(wide.flowBounds.widthPt).toBe(180);
    expect(narrow.lines.length).toBeGreaterThan(wide.lines.length);
    expect(narrow.source.path).toEqual([0]);
    expect(wide.source.path).toEqual([0]);
    expect((await renderAllPages(narrowLayout)).measures).toBe(0);
    expect((await renderAllPages(wideLayout)).measures).toBe(0);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// A5 body tables consume one retained TableLayout contract. Signed placement and
// page-local continuation ownership are acquired before paint; neither class may
// re-enter table measurement during paint.
// ─────────────────────────────────────────────────────────────────────────────

function cellOf(content: unknown[], widthPt: number): Record<string, unknown> {
  return {
    content, colSpan: 1, vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background: null, vAlign: 'top', widthPt,
  };
}

function tableOf(
  cellContent: unknown[],
  overrides: {
    colWidths?: number[];
    widthPt?: number;
    tblInd?: number;
    jc?: string;
    layout?: string;
  } = {},
): BodyElement {
  const colWidths = overrides.colWidths ?? [180];
  const widthPt = overrides.widthPt ?? colWidths[0] ?? 180;
  return {
    type: 'table',
    colWidths,
    rows: [{ cells: [cellOf(cellContent, widthPt)], rowHeight: null, rowHeightRule: 'auto', isHeader: false }],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: overrides.jc ?? 'left', layout: overrides.layout ?? 'fixed',
    ...(overrides.tblInd == null ? {} : { tblInd: overrides.tblInd }),
  } as unknown as BodyElement;
}

describe('body table retained layout paint', () => {
  it('retains page-local continuation fragments and paints them without measuring', async () => {
    const cellPara = para(Array.from({ length: 40 }, () => 'wrap').join(' '));
    const model = doc([tableOf([cellPara])], 60);
    const layout = layoutDocument(model, createLayoutServices(model, {
      localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]),
    }), { currentDateMs: 0 });
    const sliceTables = layout.pages.flatMap((page) =>
      page.layers.body.filter((node) => node.kind === 'table')) as TableFragmentLayout[];

    expect(sliceTables.length).toBeGreaterThan(1);
    const fragments: TableFragmentLayout[] = [];
    for (const fragment of sliceTables) {
      expect(fragment.rows.length).toBeGreaterThan(0);
      expect(fragment.rows.every((row) => row.occurrenceId.length > 0)).toBe(true);
      fragments.push(fragment);
    }
    expect(new Set(fragments.flatMap((fragment) =>
      fragment.rows.map((row) => row.occurrenceId))).size).toBeGreaterThan(1);

    const production = await renderAllPages(layout);
    expect(production.measures).toBe(0);
    expect(production.perPage.flat().some((call) => call.text.includes('wrap'))).toBe(true);
  });
});

describe('table-cell retained paragraph paint', () => {
  it('negative leading tblInd retains its signed origin and paints without measuring', async () => {
    const cellPara = para(Array.from({ length: 50 }, () => 'wide').join(' '));
    const model = doc([
      tableOf([cellPara], {
        colWidths: [220],
        widthPt: 220,
        tblInd: -30,
        jc: 'left',
        layout: 'fixed',
      }),
    ], 400);

    const layout = layoutDocument(model, createLayoutServices(model, {
      localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]),
    }), { currentDateMs: 0 });
    const table = layout.pages[0].layers.body[0];
    if (table?.kind !== 'table') {
      throw new Error('expected retained negative-indent table geometry');
    }
    expect(table.flowBounds.xPt).toBeCloseTo(model.section.marginLeft - 30, 6);

    const r = await assertRetainedPaint(model);
    expect(r.drawn).toBeGreaterThan(0);
    expect(r.measures).toBe(0);
  });

  it('numbered paragraph inside a table cell retains marker/body geometry', async () => {
    const numbering = { numId: 1, level: 0, format: 'decimal', text: '1.',
      indentLeft: 36, tab: 36, suff: 'tab', jc: 'left' } as unknown as DocParagraph['numbering'];
    const cellPara = {
      type: 'paragraph',
      ...para(Array.from({ length: 30 }, () => 'w').join(' '), {
        numbering, indentLeft: 36, indentFirst: -18,
      }),
    };
    const model = doc([tableOf([cellPara])], 400);
    const layout = layoutDocument(model, createLayoutServices(model, {
      localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]),
    }), { currentDateMs: 0 });
    const production = await renderAllPages(layout);
    expect(production.measures).toBe(0);
    const calls = production.perPage.flat();
    const marker = calls.find((call) => call.text === '1.');
    const body = calls.find((call) => call.text === 'w ');
    expect(marker).toBeDefined();
    expect(body).toBeDefined();
    expect(body!.x).toBeGreaterThan(marker!.x);
  });

  it('NUMPAGES inside a table cell uses the converged retained value', async () => {
    const fieldPara = {
      type: 'paragraph',
      ...para('total pages:'),
    } as unknown as DocParagraph;
    (fieldPara.runs as unknown[]).push({
      type: 'field', fieldType: 'numPages', instruction: 'NUMPAGES', fallbackText: '?',
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', background: null,
    });
    // A long body paragraph AFTER the table forces a second page, so the real
    // totalPages (2) differs from the measure-time value (1).
    const filler = para(Array.from({ length: 120 }, () => 'w').join(' '));
    const model = doc([tableOf([fieldPara]), filler as unknown as BodyElement], 100);
    const r = await assertRetainedPaint(model);
    expect(r.pages).toBeGreaterThan(1);
    expect(r.measures).toBe(0);
    // The REAL page count was drawn (not the frozen measure-time "1").
    expect(r.streams.some((page) => page.some((c) => c.text === String(r.pages)))).toBe(true);
  });

  it('acquires a later source-row PAGE field from its destination table slice', async () => {
    const pageField = para('');
    (pageField.runs as unknown[]).push({
      type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '?',
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', background: null,
    });
    const table = tableOf([para('first')]) as unknown as DocTable;
    const row = (paragraph: DocParagraph): DocTable['rows'][number] => ({
      cells: [cellOf([paragraph], 180) as unknown as DocTable['rows'][number]['cells'][number]],
      rowHeight: 24,
      rowHeightRule: 'exact',
      isHeader: false,
    });
    table.rows = [row(para('first')), row(para('second')), row(pageField)];
    const model = doc([table as unknown as BodyElement], 60);
    const layout = layoutDocument(model, createLayoutServices(model, {
      localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]),
    }), { currentDateMs: 0 });

    const pageFieldFragments = layout.pages.flatMap((page, pageIndex) =>
      page.layers.body.flatMap((node) => {
        if (node.kind === 'table') {
          return node.rows.flatMap((tableRow) =>
            tableRow.cells.flatMap((cell) =>
              cell.blocks.flatMap((block) =>
                block.layout.kind === 'paragraph' && block.layout.lines.some((line) =>
                  line.placements.some((placement) =>
                    placement.kind === 'text' && placement.dependency === 'page'))
                  ? [{ pageIndex, fragment: block.layout }]
                  : [])));
        }
        return [];
      }));
    expect(pageFieldFragments).toHaveLength(1);
    const [{ pageIndex, fragment }] = pageFieldFragments;
    expect(pageIndex).toBeGreaterThan(0);
    expect(fragment.source.path).toEqual([0, 2, 0, 0]);
    const retainedPage = fragment.lines.flatMap((line) => line.placements)
      .find((placement) => placement.kind === 'text' && placement.dependency === 'page');
    expect(retainedPage).toMatchObject({ text: String(pageIndex + 1) });

    const production = await renderAllPages(layout);
    expect(production.measures).toBe(0);
    expect(production.perPage[pageIndex].some((call) =>
      call.text === String(pageIndex + 1))).toBe(true);
  });

  it('acquires every repeated-header PAGE occurrence from its destination page', async () => {
    const pageField = para('');
    (pageField.runs as unknown[]).push({
      type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '?',
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', background: null,
    });
    const table = tableOf([pageField]) as unknown as DocTable;
    const row = (paragraph: DocParagraph, isHeader = false): DocTable['rows'][number] => ({
      cells: [cellOf([paragraph], 180) as unknown as DocTable['rows'][number]['cells'][number]],
      rowHeight: 16,
      rowHeightRule: 'exact',
      isHeader,
    });
    table.rows = [
      row(pageField, true),
      row(para('body-1')),
      row(para('body-2')),
      row(para('body-3')),
    ];
    const model = doc([table as unknown as BodyElement], 60);
    model.section = {
      ...model.section,
      pageNumType: { start: 9, fmt: 'upperRoman' },
    };
    const layout = layoutDocument(model, createLayoutServices(model, {
      localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]),
    }), { currentDateMs: 0 });

    expect(layout.pages).toHaveLength(3);
    const retainedHeaderPages = layout.pages.map((page, pageIndex) => {
      const tableNode = page.layers.body.find((node) => node.kind === 'table') as
        | TableFragmentLayout
        | undefined;
      if (tableNode?.kind !== 'table') {
        throw new Error('expected a retained table fragment');
      }
      const header = tableNode.rows.find((candidate) => candidate.logicalRowIndex === 0);
      const field = header?.cells[0]?.blocks[0]?.layout;
      const result = field?.kind === 'paragraph'
        ? field.lines.flatMap((line) => line.placements).find((placement) =>
            placement.kind === 'text' && placement.dependency === 'page')
        : undefined;
      return {
        pageIndex,
        ownership: header?.ownership,
        occurrenceId: header?.occurrenceId,
        text: result?.kind === 'text' ? result.text : undefined,
      };
    });
    expect(retainedHeaderPages).toEqual([
      { pageIndex: 0, ownership: 'source', occurrenceId: expect.any(String), text: 'IX' },
      { pageIndex: 1, ownership: 'repeated-header', occurrenceId: expect.any(String), text: 'X' },
      { pageIndex: 2, ownership: 'repeated-header', occurrenceId: expect.any(String), text: 'XI' },
    ]);

    const production = await renderAllPages(layout);
    expect(production.measures).toBe(0);
    expect(production.perPage.map((calls, pageIndex) => calls.some((call) =>
      call.text === ['IX', 'X', 'XI'][pageIndex]))).toEqual([true, true, true]);
  });

  it('acquires the following cell paragraph in the preceding float wrap context', async () => {
    // Cell paragraph 1 contributes an anchored square-wrap exclusion to the cell's
    // isolated flow domain; paragraph 2 retains the resulting wrapped partition.
    const imgPara = {
      type: 'paragraph',
      ...para(''),
    } as unknown as DocParagraph;
    (imgPara.runs as unknown[]).push({
      type: 'image',
      imagePath: 'word/media/test1.png', mimeType: 'image/png',
      widthPt: 80, heightPt: 40,
      anchor: true, anchorXPt: 100, anchorYPt: 0,
      anchorXFromMargin: false, anchorYFromPara: true,
      wrapMode: 'square', wrapSide: 'bothSides',
      anchorXRelativeFrom: 'column', anchorYRelativeFrom: 'paragraph',
    });
    const wrapPara = {
      type: 'paragraph',
      ...para(Array.from({ length: 40 }, () => 'w').join(' ')),
    };
    const model = doc([tableOf([imgPara, wrapPara])], 400);
    const r = await assertRetainedPaint(model);
    expect(r.drawn).toBeGreaterThan(0);
    expect(r.measures).toBe(0);
  });
});

describe('re-wrapped retained continuation slices (issue #908)', () => {
  it('does not re-apply first-line indent while painting a continuation', async () => {
    const p = para('あ'.repeat(28), { defaultFontSize: 20, indentFirst: 20 });
    (p.runs[0] as { fontSize: number }).fontSize = 20;
    const model = doc([p as unknown as BodyElement], 60);
    (model.section as SectionProps).columns = {
      count: 2, spacePt: 32, equalWidth: false, sep: false,
      cols: [{ widthPt: 100, spacePt: 32 }, { widthPt: 48, spacePt: 0 }],
    } as SectionProps['columns'];

    const layout = layoutDocument(model, createLayoutServices(model), { currentDateMs: 0 });
    const continuationPage = layout.pages.findIndex((page) => page.layers.body.some((node) =>
      columnIndex(page, node) === 1
        && node.kind === 'paragraph'
        && node.continuation?.continuesFromPrevious));
    expect(continuationPage).toBeGreaterThanOrEqual(0);

    const painted = await renderAllPages(layout);
    expect(painted.measures).toBe(0);
    const col1Origin = model.section.marginLeft + 100 + 32;
    const lineCalls = painted.perPage[continuationPage]
      .filter((call) => call.op === 'fill' && call.text.includes('あ') && call.x >= col1Origin);
    const callsByY = new Map<number, Call[]>();
    for (const call of lineCalls) {
      const calls = callsByY.get(call.y);
      if (calls) calls.push(call);
      else callsByY.set(call.y, [call]);
    }
    const lineStarts = [...callsByY.entries()]
      .sort(([a], [b]) => a - b)
      .map(([, calls]) => Math.min(...calls.map((call) => call.x)));

    expect(lineStarts.length).toBeGreaterThan(1);
    expect(lineStarts.slice(1).every((x) => x === lineStarts[0])).toBe(true);
  });

  it('draws a paragraph anchor once for a re-wrapped continuation', async () => {
    // A §17.6.4 unequal-width split: col0 keeps the wide partition, col1 gets a
    // RE-MEASURED remainder whose fragment covers its WHOLE partition (lineStart 0,
    // lineEnd = length). paintParagraphFragment must not degrade that full-range
    // continuation slice to "no slice": renderParagraph's first-slice-only work
    // (anchor drawing, §17.3.1.7 top border) would run again in the destination
    // column. The retained continuation owns the source occurrence only once. The
    // anchored no-wrap shape's text is the observable.
    const anchorShape = {
      type: 'shape',
      widthPt: 40, heightPt: 10,
      anchorXPt: 0, anchorYPt: 0,
      anchorXFromMargin: false, anchorYFromPara: true,
      anchorXRelativeFrom: 'column', anchorYRelativeFrom: 'paragraph',
      zOrder: 0, subpaths: [], presetGeometry: 'rect',
      fill: null, stroke: null,
      // No wrap: the shape registers no float band, so the width-mismatch swap
      // fires (a wrap float would scope the swap out — asserted below).
      wrapMode: 'none', wrapSide: 'bothSides',
      distTop: 0, distBottom: 0, distLeft: 0, distRight: 0,
      textBlocks: [{
        text: 'ANCHORTEXT', fontSizePt: 8, bold: false, italic: false,
        color: '000000', fontFamily: 'Times New Roman', alignment: 'left',
      }],
      textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
    };
    const p = para('あ'.repeat(28), { defaultFontSize: 20 });
    (p.runs as unknown[]).unshift(anchorShape);
    (p.runs[1] as { fontSize: number }).fontSize = 20;
    const model = doc([p as unknown as BodyElement], 60);
    (model.section as SectionProps).columns = {
      count: 2, spacePt: 32, equalWidth: false, sep: false,
      cols: [{ widthPt: 100, spacePt: 32 }, { widthPt: 48, spacePt: 0 }],
    } as SectionProps['columns'];

    const layout = layoutDocument(model, createLayoutServices(model), { currentDateMs: 0 });
    // Non-vacuity: the swap really fired — a col-1 slice carries the remainder
    // partition marker (if a float had registered, the swap would be scoped out
    // and this test would be exercising nothing).
    const slices = layout.pages.flatMap((page) =>
      page.layers.body.filter((node) => node.kind === 'paragraph'));
    expect(slices.some((node) => node.continuation?.continuesFromPrevious)).toBe(true);

    const production = await renderAllPages(layout);
    expect(production.measures).toBe(0);
    for (let pg = 0; pg < production.perPage.length; pg++) {
      const anchorDrawsProduction = production.perPage[pg].filter((c) => c.text === 'ANCHORTEXT').length;
      expect(anchorDrawsProduction).toBe(1);
    }
    expect(production.perPage.flat().some((c) => c.text === 'ANCHORTEXT')).toBe(true);
  });

  it('draws behindDoc paragraph anchors once for a re-wrapped continuation', async () => {
    const anchorShape = {
      type: 'shape',
      widthPt: 40, heightPt: 10,
      anchorXPt: 0, anchorYPt: 0,
      anchorXFromMargin: false, anchorYFromPara: true,
      anchorXRelativeFrom: 'column', anchorYRelativeFrom: 'paragraph',
      zOrder: 0, behindDoc: true,
      subpaths: [], presetGeometry: 'rect',
      fill: null, stroke: null,
      wrapMode: 'none', wrapSide: 'bothSides',
      distTop: 0, distBottom: 0, distLeft: 0, distRight: 0,
      textBlocks: [{
        text: 'ANCHORTEXT', fontSizePt: 8, bold: false, italic: false,
        color: '000000', fontFamily: 'Times New Roman', alignment: 'left',
      }],
      textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
    };
    const p = para('あ'.repeat(28), { defaultFontSize: 20 });
    (p.runs as unknown[]).unshift(anchorShape);
    (p.runs[1] as { fontSize: number }).fontSize = 20;
    const model = doc([p as unknown as BodyElement], 60);
    (model.section as SectionProps).columns = {
      count: 2, spacePt: 32, equalWidth: false, sep: false,
      cols: [{ widthPt: 100, spacePt: 32 }, { widthPt: 48, spacePt: 0 }],
    } as SectionProps['columns'];

    const layout = layoutDocument(model, createLayoutServices(model), { currentDateMs: 0 });
    const slices = layout.pages.flatMap((page) =>
      page.layers.body.filter((node) => node.kind === 'paragraph'));
    expect(slices.some((node) => node.continuation?.continuesFromPrevious)).toBe(true);

    const production = await renderAllPages(layout);
    expect(production.measures).toBe(0);
    for (let pg = 0; pg < production.perPage.length; pg++) {
      const anchorDrawsProduction = production.perPage[pg].filter((c) => c.text === 'ANCHORTEXT').length;
      expect.soft(anchorDrawsProduction).toBe(1);
    }
  });
});
