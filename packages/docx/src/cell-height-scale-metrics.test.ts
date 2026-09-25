import { beforeAll, describe, it, expect } from 'vitest';
import { layoutDocument } from './document-layout.js';
import { paintLayoutPage } from './paint/canvas-page.js';
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  DocxTextRun,
  SectionProps,
} from './types';

// ─────────────────────────────────────────────────────────────────────────────
// Cell height measurement and glyph paint share canonical scale-1 geometry.
//
// Pagination resolves ordinary body-table text in document coordinates. Paint
// must preserve that scale-1 line geometry and map its glyphs through a Canvas
// viewport transform; cell measurement must scale the same canonical line boxes.
// If either side instead asks Canvas for fresh paint-size metrics, hinted fonts
// can make row height / vAlign disagree with the glyph box actually drawn.
//
// These tests mock a NON-LINEAR `measureText` — a sub-linear ascent/descent, the
// direction real font hinting bends (glyphs proportionally shorter at larger
// sizes) — so `metric(12·s) ≠ s·metric(12)`. They then render at scale = 4/3
// (≈1.333, the renderer's default cssWidth/physWidth ratio) and assert the cell
// content height that vAlign / row layout USES equals the transformed glyph box
// the paint side actually PRODUCES, read back through the public render seam.
// ─────────────────────────────────────────────────────────────────────────────

/** Sub-linear single-glyph vertical metrics in px at font size `p` px. The
 *  ascent/descent shrink proportionally as `p` grows, so `metric(p·s) ≠ s·metric(p)`
 *  — a hinting-like non-linearity that would expose a paint-size remeasurement. Width is
 *  kept linear (charCount · p · 0.5) so the line PARTITION never changes with
 *  scale; only the native paint-size BOX height is scale-non-linear. */
function glyphMetrics(p: number): { asc: number; desc: number } {
  return { asc: p * (0.8 - 0.01 * p), desc: p * (0.2 - 0.005 * p) };
}

interface FillTextCall { text: string; x: number; y: number; font: string; scaleY: number; }
interface FillRectCall { x: number; y: number; w: number; h: number; fillStyle: string; }

function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  fillTextCalls: FillTextCall[];
  fillRectCalls: FillRectCall[];
  measured: () => number;
} {
  let font = '10px serif';
  let fillStyle = '#000';
  let transform = { scaleX: 1, scaleY: 1, translateX: 0, translateY: 0 };
  const stack: typeof transform[] = [];
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const fillTextCalls: FillTextCall[] = [];
  const fillRectCalls: FillRectCall[] = [];
  let measured = 0;
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get fillStyle() { return fillStyle; },
    set fillStyle(v: string) { fillStyle = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      measured += 1;
      const p = px();
      const { asc, desc } = glyphMetrics(p);
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: asc,
        fontBoundingBoxDescent: desc,
        actualBoundingBoxAscent: asc,
        actualBoundingBoxDescent: desc,
      } as TextMetrics;
    },
    save() { stack.push({ ...transform }); },
    restore() { transform = stack.pop() ?? transform; },
    setTransform(a: number, _b: number, _c: number, d: number, e: number, f: number) {
      transform = { scaleX: a, scaleY: d, translateX: e, translateY: f };
    },
    beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {},
    fillRect(x: number, y: number, w: number, h: number) {
      fillRectCalls.push({
        x: transform.translateX + transform.scaleX * x,
        y: transform.translateY + transform.scaleY * y,
        w: transform.scaleX * w,
        h: transform.scaleY * h,
        fillStyle,
      });
    },
    strokeRect() {}, clip() {}, rect() {},
    scale(x: number, y: number) {
      transform.scaleX *= x;
      transform.scaleY *= y;
    },
    translate(x: number, y: number) {
      transform.translateX += transform.scaleX * x;
      transform.translateY += transform.scaleY * y;
    },
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText(text: string, x: number, y: number) {
      fillTextCalls.push({
        text,
        x: transform.translateX + transform.scaleX * x,
        y: transform.translateY + transform.scaleY * y,
        font,
        scaleY: transform.scaleY,
      });
    },
    strokeText() {},
    strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0, height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  // `clipExact` rows (ST_HeightRule exact) read `ctx.canvas.width`; wire the
  // back-reference so the exact-height cells below render.
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return {
    canvas: canvas as unknown as HTMLCanvasElement,
    fillTextCalls,
    fillRectCalls,
    measured: () => measured,
  };
}

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    getContext() { return makeRecordingCanvas().canvas.getContext('2d'); }
  };
});

// An uncatalogued synthetic font has no reference single-line floor; the line
// box is exactly the mock's ascent +
// descent — isolating the metric non-linearity that the canonical path must avoid.
const TEST_FONT = 'Synthetic Untabled Serif';

function textRun(text: string): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 12, color: null, fontFamily: TEST_FONT, fontFamilyEastAsia: '',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  };
}

function paraOf(text: string, opts: Partial<DocParagraph> = {}): CellElement {
  return {
    type: 'paragraph',
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{ type: 'text', ...textRun(text) } as DocParagraph['runs'][number]],
    defaultFontSize: 12, defaultFontFamily: TEST_FONT,
    widowControl: false,
    ...opts,
  } as unknown as CellElement;
}

function cell(
  content: CellElement[],
  vAlign: 'top' | 'center' | 'bottom',
  background: string | null = null,
): DocTableCell {
  return {
    content,
    colSpan: 1,
    vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background,
    vAlign,
    widthPt: 300,
  } as DocTableCell;
}

function row(c: DocTableCell, rowHeight: number | null, rule: 'auto' | 'atLeast' | 'exact'): DocTableRow {
  return {
    cells: [c],
    rowHeight,
    rowHeightRule: rule,
    isHeader: false,
  } as DocTableRow;
}

function tableOf(r: DocTableRow): DocTable {
  return {
    colWidths: [300],
    rows: [r],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
    // Negative leading indents remain represented by retained table geometry.
    tblInd: -10,
  } as DocTable;
}

function docWithTable(t: DocTable): DocxDocumentModel {
  return {
    section: {
      pageWidth: 300, pageHeight: 600,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'table', ...t } as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { [TEST_FONT]: 'roman' },
  } as unknown as DocxDocumentModel;
}

// scale = cssWidth / pageWidthPt = 400 / 300 = 4/3 ≈ 1.333.
const CSS_WIDTH = 400;
const PAGE_WIDTH = 300;
const SCALE = CSS_WIDTH / PAGE_WIDTH;

async function renderAndRead(t: DocTable) {
  const model = docWithTable(t);
  const layout = layoutDocument(model);
  const table = layout.pages.flatMap((page) => page.layers.body)
    .find((node) => node.kind === 'table');
  expect(table).toBeDefined();
  if (table?.kind !== 'table') {
    throw new Error('expected retained TableLayout/TableFragmentLayout');
  }
  const rec = makeRecordingCanvas();
  await paintLayoutPage(layout, 0, rec.canvas, { dpr: 1, scale: CSS_WIDTH / PAGE_WIDTH });
  expect(rec.measured()).toBe(0);
  return rec;
}

/** Painted inked-block extent (px) of a set of one-line paragraphs, read from the
 *  transformed fillText baselines and the scale-1 glyph box mapped by its CTM. */
function paintedExtent(calls: FillTextCall[], firstText: string, lastText: string): { top: number; bottom: number } {
  const first = calls.find((c) => c.text === firstText);
  const last = calls.find((c) => c.text === lastText);
  expect(first).toBeDefined();
  expect(last).toBeDefined();
  const firstPx = parseFloat(/(\d+(?:\.\d+)?)px/.exec(first!.font)?.[1] ?? '12');
  const lastPx = parseFloat(/(\d+(?:\.\d+)?)px/.exec(last!.font)?.[1] ?? '12');
  const firstMetrics = glyphMetrics(firstPx);
  const lastMetrics = glyphMetrics(lastPx);
  return {
    top: first!.y - firstMetrics.asc * first!.scaleY,
    bottom: last!.y + lastMetrics.desc * last!.scaleY,
  };
}

describe('cell content height matches canonical transformed Canvas metrics', () => {
  it('sanity: the mock ascent/descent is scale-NON-linear (else the test is vacuous)', () => {
    const one = glyphMetrics(12);
    const s = glyphMetrics(12 * SCALE);
    const oneH = one.asc + one.desc;
    const scaledH = s.asc + s.desc;
    // Native paint-size metrics must differ clearly from the canonical scale-1
    // box × viewport scale, or an accidental paint-size remeasurement could pass.
    expect(Math.abs(oneH * SCALE - scaledH)).toBeGreaterThan(0.5);
  });

  it('vAlign=center: the inked block is centred using the PAINTED content height', async () => {
    // Three single-line paragraphs, no spacing. Exact 120 pt row → drawn height is
    // exactly 120·scale px (resolveSingleRowHeight, ST_HeightRule exact) and is
    // NOT content-measured, so this isolates the vAlign centring consumer. The
    // content (~3 lines) fits well within the row, leaving centring room.
    const t = tableOf(row(
      cell(
        [paraOf('aa'), paraOf('bb'), paraOf('cc')],
        'center',
      ),
      120,
      'exact',
    ));
    const { fillTextCalls } = await renderAndRead(t);
    const { top, bottom } = paintedExtent(fillTextCalls, 'aa', 'cc');
    const inkedMid = (top + bottom) / 2;
    const cellMid = (120 * SCALE) / 2; // row at y=0, exact height 120·scale.
    // measure == paint: vAlign used the same canonical transformed content height,
    // so the painted block's midpoint lands on the cell midpoint.
    expect(inkedMid).toBeCloseTo(cellMid, 1);
  });

  it('vAlign=bottom: the inked block bottom hugs the cell bottom (painted height)', async () => {
    const t = tableOf(row(
      cell([paraOf('aa'), paraOf('bb'), paraOf('cc')], 'bottom'),
      120,
      'exact',
    ));
    const { fillTextCalls } = await renderAndRead(t);
    const { bottom } = paintedExtent(fillTextCalls, 'aa', 'cc');
    // mb = 0 → the transformed inked bottom sits on the cell bottom (120·scale).
    expect(bottom).toBeCloseTo(120 * SCALE, 1);
  });

  it('auto row height fallback: reserved row height equals the painted content height', async () => {
    // Auto height → the retained row height is the measured cell content height.
    const t = tableOf(row(
      cell([paraOf('aa'), paraOf('bb'), paraOf('cc')], 'top', 'abcdef'),
      null,
      'auto',
    ));
    (t as unknown as DocTable).tblInd = -1;
    const { fillTextCalls, fillRectCalls } = await renderAndRead(t);
    const bg = fillRectCalls.find((r) => r.fillStyle === '#abcdef');
    expect(bg).toBeDefined();
    const { top, bottom } = paintedExtent(fillTextCalls, 'aa', 'cc');
    // Zero cell margins → the reserved row height equals the transformed glyph
    // extent sourced from the retained layout.
    expect(bg!.h).toBeCloseTo(bottom - top, 1);
  });

  it('a negative-indent vAlign table paints retained scale-1 geometry at viewport scale', async () => {
    // The §17.4.84 centring calculation and paint both consume the retained row
    // geometry, so nonlinear paint-size metrics cannot shift the drawn height.
    const t = tableOf(row(
      cell([paraOf('aa'), paraOf('bb'), paraOf('cc')], 'center', 'abcdef'),
      null,
      'auto',
    ));
    const { fillRectCalls } = await renderAndRead(t);
    const bg = fillRectCalls.find((r) => r.fillStyle === '#abcdef');
    expect(bg).toBeDefined();
    // The paginator resolved the row at scale 1: 3 one-line paragraphs at 12pt with
    // the mock's scale-1 glyph box (asc+desc at p=12), so the production-drawn row
    // height is that scale-1 height × SCALE, matching the canonical glyph transform.
    const { asc, desc } = glyphMetrics(12);
    const scale1RowHeightPt = 3 * (asc + desc);
    expect(bg!.h).toBeCloseTo(scale1RowHeightPt * SCALE, 1);
  });
});
